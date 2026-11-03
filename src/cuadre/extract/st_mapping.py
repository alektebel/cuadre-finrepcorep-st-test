"""Extracción determinista desde las plantillas oficiales ST de la EBA.

Produce, sin LLM:

1. Inventario de celdas de punto de partida (para el traductor y los checks).
2. Reglas ``plantilla`` de mapeo a COREP, a partir de la columna ``COREP CODE /
   COMPUTATION`` (y, en riesgo de crédito, de la fila de cabecera con la columna COREP).
3. Reglas ``plantilla`` de controles internos (sumas) en las columnas de punto de partida.
4. Lista de incidencias: textos de mapeo que no son celdas COREP traducibles.
"""
from __future__ import annotations

import collections
import datetime as dt
import json
from pathlib import Path

import yaml

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ..st_template import (_COLS_ALT, _ROWS_ALT, StSheet, col_label, parse_alternatives, parse_full_cells,
                           parse_sheet, row_label)
from .inventory import formula_to_rule


def _is_ratio_col(label: str) -> bool:
    lab = label.lower()
    return "%" in lab or "pd assigned" in lab


def _corep_ref(tpl: str, r: str, c: str, z: str | None) -> str:
    return f"COREP:{tpl}!r{r}c{c}" + (f"@{z}" if z else "")


def _expr(terms: list[tuple[float, str]]) -> str:
    parts = []
    for k, ref in terms:
        mag = "" if abs(k) == 1 else f"{abs(k):g}*"
        parts.append(f"{'-' if k < 0 else '+'} {mag}{{{ref}}}")
    return " ".join(parts).removeprefix("+ ")


def _classify(text: str) -> str:
    t = text.strip()
    if t.lower().startswith("projection"):
        return "proyeccion"
    if t.startswith("[") or "/" in t or t.startswith(("R.", "D.", "A.", "B.")):
        return "calculo_interno"
    if "%" in t:
        return "porcentaje"
    if "minus" in t.lower() or "(" in t and " or " in t:
        return "compuesto_no_soportado"
    return "otro"


def extract(xlsx: str | Path, prefix: str = "ST27") -> dict:
    xlsx = Path(xlsx)
    wv = load_workbook(xlsx, data_only=True)
    wf = load_workbook(xlsx, data_only=False)
    cells, rules, issues = [], [], collections.Counter()
    issue_examples: dict[str, list[str]] = collections.defaultdict(list)
    per_sheet = {}

    def issue(kind: str, where: str, text: str) -> None:
        issues[kind] += 1
        if len(issue_examples[kind]) < 5:
            issue_examples[kind].append(f"{where}: {text[:120]}")

    for ws in wv.worksheets:
        st: StSheet | None = parse_sheet(ws, wf[ws.title])
        if st is None or not st.rownums or not st.colnums:
            continue
        sf = wf[ws.title]
        n_map = n_tpl = 0
        start_rows = [r for r in st.rownums if st.is_starting_point_row(ws, r)]
        start_cols = sorted(st.start_cols) if st.start_cols else sorted(st.colnums)
        labels_c = {c: col_label(ws, st, c) for c in st.colnums}

        # inventario y fórmulas internas (solo punto de partida)
        addr2ref = {f"{get_column_letter(c)}{r}": st.ref(r, c) for r in st.rownums for c in st.colnums}
        for r in start_rows:
            rl = row_label(ws, st, r)
            for c in start_cols:
                f = sf.cell(r, c).value
                ref = st.ref(r, c)
                cells.append({"ref": str(ref), "sheet": ws.title, "cell": f"{get_column_letter(c)}{r}",
                              "row_label": rl, "col_label": labels_c[c],
                              "formula": f if isinstance(f, str) and f.startswith("=") else None})
                if isinstance(f, str) and f.startswith("="):
                    rule = formula_to_rule(str(ref), f, addr2ref)
                    if rule:
                        rule["id"] = f"{prefix}-TPL-{ws.title}-r{ref.row}c{ref.col}"
                        rule["fuente"] = {"documento": xlsx.name, "hoja": ws.title,
                                          "celda": f"{get_column_letter(c)}{r}", "formula": f}
                        rules.append(rule)
                        n_tpl += 1

        # mapeo COREP
        if st.map_col:
            col_maps = {}
            if st.colmap_row:
                for c in st.colnums:
                    h = sf.cell(st.colmap_row, c).value
                    if isinstance(h, str) and h.strip() and not _is_ratio_col(labels_c[c]):
                        alt = parse_alternatives(h, _COLS_ALT)
                        if alt:
                            col_maps[c] = (alt, h.strip())
                        else:
                            issue("columna_compuesta_no_soportada", f"{ws.title}!{get_column_letter(c)}{st.colmap_row}", h)
            for r in st.rownums:
                text = sf.cell(r, st.map_col).value
                if not isinstance(text, str) or text.strip() in ("", "-"):
                    continue
                where = f"{ws.title}!{get_column_letter(st.map_col)}{r}"
                if text.strip().lower().startswith("projection") or not st.is_starting_point_row(ws, r):
                    issue("proyeccion", where, text)
                    continue
                rl = row_label(ws, st, r)
                geo_col = st.attr_cols.get("geo")
                geo = ws.cell(r, geo_col).value if geo_col else None
                z = f"GEO={str(geo).strip()}" if geo not in (None, "") else None
                full = parse_full_cells(text)
                if full is not None:
                    for c in start_cols:
                        ref = st.ref(r, c)
                        rhs = _expr([(k, _corep_ref(t, rr, cc, None)) for k, t, rr, cc in full])
                        rules.append({
                            "id": f"{prefix}-MAP-{ws.title}-r{ref.row}c{ref.col}",
                            "descripcion": f"{rl} / {labels_c[c]} = COREP"[:200],
                            "tipo": "igualdad", "lhs": "{" + str(ref) + "}", "rhs": rhs, "origen": "plantilla",
                            "fuente": {"documento": xlsx.name, "hoja": ws.title, "celda": where.split("!")[1],
                                       "cita": text.strip()},
                            "etiquetas": ["mapeo_corep"],
                        })
                        n_map += 1
                    continue
                rows_alt = parse_alternatives(text, _ROWS_ALT)
                if rows_alt is None:
                    issue(_classify(text), where, text)
                    continue
                if not col_maps:
                    issue("fila_sin_columnas_corep", where, text)
                    continue
                for c, (cols_alt, htext) in col_maps.items():
                    tpls = [t for t in rows_alt if t in cols_alt]
                    if not tpls:
                        issue("plantilla_fila_y_columna_distintas", where, f"{text} × {htext}")
                        continue
                    ref = st.ref(r, c)
                    for t in tpls:
                        terms = [(1.0, _corep_ref(t, rr, cc, z)) for rr in rows_alt[t] for cc in cols_alt[t]]
                        suffix = f"-{t.replace(' ', '').replace('.', '')}" if len(tpls) > 1 else ""
                        rules.append({
                            "id": f"{prefix}-MAP-{ws.title}-r{ref.row}c{ref.col}{suffix}",
                            "descripcion": f"{rl} / {labels_c[c]} = COREP {t}"[:200],
                            "tipo": "igualdad", "lhs": "{" + str(ref) + "}", "rhs": _expr(terms),
                            "origen": "plantilla",
                            "fuente": {"documento": xlsx.name, "hoja": ws.title,
                                       "celda": f"{where.split('!')[1]} × {get_column_letter(c)}{st.colmap_row}",
                                       "cita": f"{text.strip()} | {htext}"},
                            "etiquetas": ["mapeo_corep"] + (["alternativa"] if len(tpls) > 1 else []),
                        })
                        n_map += 1
        per_sheet[ws.title] = {"filas": len(st.rownums), "columnas": len(st.colnums),
                               "filas_punto_partida": len(start_rows), "cols_punto_partida": len(start_cols),
                               "reglas_mapeo": n_map, "reglas_formula": n_tpl, "avisos": st.warnings}
    return {
        "inventory": {"file": str(xlsx), "framework": "ST", "cells": cells, "template_rules": [],
                      "issues": [f"{k}: {v}" for k, v in issues.items()]},
        "rules": rules,
        "issues": dict(issues),
        "issue_examples": dict(issue_examples),
        "sheets": per_sheet,
    }


def write_outputs(res: dict, rules_dir: Path, work: Path, internal_controls: bool = False) -> str:
    """Escribe inventario (work/), reglas (rules_dir/) e informe de extracción."""
    rules_dir.mkdir(parents=True, exist_ok=True)
    work.mkdir(parents=True, exist_ok=True)
    stamp = dt.date.today().isoformat()
    src = Path(res["inventory"]["file"]).name
    (work / "inventory_st.json").write_text(json.dumps(res["inventory"], ensure_ascii=False), encoding="utf-8")
    # Resumen compacto para agentes: filas y columnas por hoja, no su producto.
    compact: dict = {}
    for c in res["inventory"]["cells"]:
        sh = compact.setdefault(c["sheet"], {"filas": {}, "columnas": {}})
        ref = c["ref"].split("!")[1]
        row, col = ref[1:5], ref[6:10]
        sh["filas"].setdefault(row, c["row_label"])
        sh["columnas"].setdefault(col, c["col_label"])
    (work / "resumen_inventario_st.json").write_text(
        json.dumps({"fichero": src, "formato_ref": "ST:<hoja>!r<RowNum>c<col>", "hojas": compact},
                   ensure_ascii=False, indent=1), encoding="utf-8")

    mapping = [r for r in res["rules"] if "-MAP-" in r["id"]]
    internal = [r for r in res["rules"] if "-TPL-" in r["id"]]
    head = (f"# Generado por 'cuadre-extract st-plantillas' el {stamp} desde '{src}'.\n"
            "# Extracción determinista (sin LLM). No editar a mano: regenerar.\n")
    (rules_dir / "plantilla_mapeo_corep.yaml").write_text(
        head + "# Cuadres de punto de partida ST = COREP según la columna 'COREP CODE / COMPUTATION'.\n"
        + yaml.safe_dump({"reglas": mapping}, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")
    if internal_controls:
        (rules_dir / "plantilla_controles_internos.yaml.opcional").write_text(
            head + "# Fórmulas internas de la plantilla en columnas de punto de partida. Extensión '.opcional'\n"
            "# para que el ejercicio no las cargue por defecto; renombrar a .yaml para activarlas.\n"
            + yaml.safe_dump({"reglas": internal}, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")

    by_sheet = collections.Counter(r["fuente"]["hoja"] for r in mapping)
    alt = sum(1 for r in mapping if "alternativa" in r.get("etiquetas", []))
    lines = [f"# Extracción desde plantillas ST: {src}", "", f"Fecha: {stamp}", "",
             f"- Reglas de mapeo a COREP: **{len(mapping)}** ({alt} son alternativas C 09.02 / C 09.05)",
             f"- Controles internos (fórmulas) en punto de partida: {len(internal)}"
             + (" (exportados como .opcional)" if internal_controls else " (no exportados)"),
             f"- Celdas de punto de partida inventariadas: {len(res['inventory']['cells'])}", "",
             "## Reglas de mapeo por hoja", "", "| Hoja | Reglas |", "|---|---|"]
    lines += [f"| {k} | {v} |" for k, v in sorted(by_sheet.items())]
    lines += ["", "## Textos de mapeo no convertidos", "", "| Tipo | Nº | Ejemplos |", "|---|---|---|"]
    for k, v in sorted(res["issues"].items()):
        ex = "<br>".join(e.replace("|", "\\|") for e in res["issue_examples"].get(k, [])[:3])
        lines.append(f"| {k} | {v} | {ex} |")
    lines += ["", "## Hojas", "", "| Hoja | Filas | Columnas | Filas punto partida | Cols punto partida | Avisos |",
              "|---|---|---|---|---|---|"]
    for k, d in res["sheets"].items():
        lines.append(f"| {k} | {d['filas']} | {d['columnas']} | {d['filas_punto_partida']} | "
                     f"{d['cols_punto_partida']} | {'; '.join(d['avisos'])} |")
    (rules_dir / "INFORME_EXTRACCION.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return f"{len(mapping)} reglas de mapeo, {len(internal)} controles internos -> {rules_dir}"
