"""Genera el libro Excel autónomo de cuadre (fórmulas vivas, sin Python en el puesto).

Fuente única de verdad: el catálogo YAML (reglas + conceptos + parámetros de gravedad).
Este módulo lo "compila" a un libro cuyas fórmulas reproducen el motor Python:

    Datos      (un hecho por fila; lo rellena la macro de carga o --precargar)
       ▲ MATCH/INDEX por clave
    Operandos  (un término de regla por fila: clave, valor, redondeo, origen, enlaces)
       ▲ SUMIFS/COUNTIFS por regla
    Checks     (una regla por fila: LHS, RHS, diferencia, ruido, estado, gravedad)
       ▲
    Resumen

Las fórmulas de las columnas calculadas de ``Datos`` se guardan en notación R1C1 en la
hoja oculta ``_Formulas``: la macro VBA las lee de ahí al cargar ficheros, de modo que
no hay una segunda copia de esas fórmulas en el código VBA.

Compatibilidad: solo funciones disponibles desde Excel 2010 (SUMIFS, COUNTIFS, INDEX,
MATCH, IFERROR, HYPERLINK). Sin matrices dinámicas ni XLOOKUP.
"""
from __future__ import annotations

import datetime as dt
import os
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from .. import __version__
from ..catalog import ConceptMap, load_adjustments, load_rules
from ..expressions import parse_expression
from ..loaders import load_source
from ..run import Exercise, load_sources, sha256
from ..severity import SeverityConfig

MAX_DATA_ROWS = 200_000
MAX_ADJ_ROWS = 2_000
MAX_LOG_ROWS = 5_000

COLORS = {
    "Crítico": "C00000", "Grave": "FF4D4D", "Moderado": "FFA64D", "Leve": "FFE066",
    "REDONDEO": "BDD7EE", "OK": "C6EFCE", "NO_EVALUABLE": "D9D9D9", "KO": "FF4D4D",
}
HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
INPUT_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
EUR_FMT = '#,##0.00;[Red]-#,##0.00'

# --- Datos: columnas y fórmulas (R1C1, solo referencias de la misma fila RC<n>) -------
DATOS_COLS = [
    ("Clave", 40), ("Marco", 9), ("Plantilla", 16), ("Fila", 7), ("Columna", 8), ("Z", 12),
    ("Valor_fichero", 18), ("Escala", 10), ("Valor_EUR", 20), ("Decimales", 10),
    ("Redondeo_EUR", 14), ("Fichero", 50), ("Hoja", 18), ("Celda", 8), ("Abrir", 8),
    ("Implicado", 11), ("Duplicado", 11),
]
DATOS_FORMULAS_R1C1 = {
    1: '=RC2&":"&RC3&"!r"&RC4&"c"&RC5&IF(RC6="","","@"&RC6)',
    9: "=RC7*RC8",
    11: '=IF(RC10="",0,0.5*RC8*10^(-RC10))',
    15: '=IF(RC12="","",HYPERLINK(RC12&"#\'"&RC13&"\'!"&RC14,"abrir"))',
    16: '=IF(COUNTIFS(O_Clave,RC1,O_Estado,"KO")>0,"KO",IF(COUNTIFS(O_Clave,RC1,O_Estado,"REDONDEO")>0,"REDONDEO",""))',
    17: '=IF(COUNTIF(D_Clave,RC1)>1,"DUPLICADO","")',
}
# Columnas que escribe la carga (valores); el resto son fórmulas.
DATOS_VALUE_COLS = [2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14]

_RC = re.compile(r"RC(\d+)")


def r1c1_to_a1(formula: str, row: int) -> str:
    """Convierte las referencias ``RC<n>`` (misma fila) a A1. No admite otras formas."""
    return _RC.sub(lambda m: f"{get_column_letter(int(m.group(1)))}{row}", formula)


def _name(wb: Workbook, name: str, ref: str) -> None:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _header(ws, cols: list[tuple[str, int]], row: int = 1) -> None:
    for i, (title, width) in enumerate(cols, start=1):
        c = ws.cell(row, i, title)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = ws.cell(row + 1, 1)


def _status_formatting(ws, rng: str) -> None:
    for status in ("KO", "NO_EVALUABLE", "REDONDEO", "OK"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{status}"'],
                                                      fill=PatternFill("solid", start_color=COLORS[status], end_color=COLORS[status])))


def _severity_formatting(ws, rng: str, levels: list[str]) -> None:
    for lvl in levels:
        col = COLORS.get(lvl, "FFFFFF")
        font = Font(color="FFFFFF", bold=True) if lvl == levels[-1] else None
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{lvl}"'],
                                                      fill=PatternFill("solid", start_color=col, end_color=col), font=font))


def _source_text(src: dict) -> str:
    parts = [str(src[k]) for k in ("documento", "apartado") if src.get(k)]
    if src.get("pagina"):
        parts.append(f"p. {src['pagina']}")
    return " · ".join(parts)


def build_workbook(exercise_path: str | Path, out: str | Path, preload_sources: str | Path | None = None) -> Path:
    ex = Exercise.load(exercise_path)
    cmap = ConceptMap.load(ex.concept_files)
    rules = [r for r in load_rules(ex.rule_patterns) if r.enabled]
    sev = SeverityConfig.load(ex.severity_file)
    if len(sev.levels) != 4:
        raise ValueError("El libro Excel asume 4 grados de gravedad")

    wb = Workbook()
    ws_inicio = wb.active
    ws_inicio.title = "Inicio"
    ws_res = wb.create_sheet("Resumen")
    ws_chk = wb.create_sheet("Checks")
    ws_ops = wb.create_sheet("Operandos")
    ws_dat = wb.create_sheet("Datos")
    ws_log = wb.create_sheet("Carga")
    ws_map = wb.create_sheet("Mapeo_hojas")
    ws_con = wb.create_sheet("Conceptos")
    ws_adj = wb.create_sheet("Ajustes")
    ws_par = wb.create_sheet("Parametros")
    ws_meta = wb.create_sheet("Metadatos")
    ws_f = wb.create_sheet("_Formulas")
    ws_f.sheet_state = "hidden"

    # ------------------------------------------------------------- Parametros
    ws_par["A1"] = "Parámetros del cuadre (celdas amarillas editables)"
    ws_par["A1"].font = Font(bold=True, size=13)
    ws_par["A3"], ws_par["B3"] = "Tolerancia (EUR)", sev.tolerance_eur
    ws_par["A5"], ws_par["B5"], ws_par["C5"], ws_par["D5"] = "Rango", "Gravedad", "Límite absoluto (EUR, <)", "Límite relativo (<)"
    for i, lvl in enumerate(sev.levels):
        r = 6 + i
        ws_par.cell(r, 1, i + 1)
        ws_par.cell(r, 2, lvl)
        if i < 3:
            ws_par.cell(r, 3, sev.abs_thresholds_eur[i]).fill = INPUT_FILL
            ws_par.cell(r, 4, sev.rel_thresholds[i]).fill = INPUT_FILL
            ws_par.cell(r, 4).number_format = "0.00%"
    ws_par["B3"].fill = INPUT_FILL
    ws_par["A12"], ws_par["B12"], ws_par["C12"] = "Marco", "Escala por defecto", "Decimales por defecto"
    defaults = [("ST", 1_000_000, 2), ("FINREP", 1, 2), ("COREP", 1, 2)]
    for i, (m, e, d) in enumerate(defaults):
        ws_par.cell(13 + i, 1, m)
        ws_par.cell(13 + i, 2, e).fill = INPUT_FILL
        ws_par.cell(13 + i, 3, d).fill = INPUT_FILL
    ws_par["A17"] = ("Escala: multiplicador a EUR (1.000.000 si el fichero va en millones). Decimales: precisión con la "
                     "que se reporta; define el redondeo admitido de cada dato (0,5 × escala × 10^-decimales). "
                     "VERIFICAR en la Template Guidance la unidad de las plantillas ST.")
    ws_par["A17"].alignment = Alignment(wrap_text=True, vertical="top")
    ws_par.merge_cells("A17:D20")
    ws_par["A22"], ws_par["B22"] = "Cargar solo datos usados por reglas (SI/NO)", "SI"
    ws_par["B22"].fill = INPUT_FILL
    ws_par["A23"] = ("Con SI la macro omite las celdas que ninguna regla usa: una plantilla ST completa "
                     "puede superar el máximo de filas de Datos.")
    for col, w in zip("ABCD", (22, 20, 24, 20)):
        ws_par.column_dimensions[col].width = w
    _name(wb, "P_Tol", "Parametros!$B$3")
    _name(wb, "P_Niveles", "Parametros!$B$6:$B$9")
    _name(wb, "P_Marcos", "Parametros!$A$13:$C$15")
    _name(wb, "P_SoloUsados", "Parametros!$B$22")
    for i in range(3):
        _name(wb, f"P_Abs{i + 1}", f"Parametros!$C${6 + i}")
        _name(wb, f"P_Rel{i + 1}", f"Parametros!$D${6 + i}")

    # --------------------------------------------------------------- Conceptos
    _header(ws_con, [("Concepto", 34), ("Clave", 44), ("Etiqueta", 34), ("Verificado", 11), ("Fichero", 40)])
    for i, c in enumerate(sorted(cmap.concepts.values(), key=lambda c: c.name), start=2):
        ws_con.cell(i, 1, c.name)
        ws_con.cell(i, 2, str(c.ref)).fill = INPUT_FILL
        ws_con.cell(i, 3, c.label)
        ws_con.cell(i, 4, "sí" if c.verified else "no")
        ws_con.cell(i, 5, Path(c.origin_file or "").name)
    n_con = max(2, len(cmap.concepts) + 1)
    _name(wb, "K_Tabla", f"Conceptos!$A$2:$B${n_con + 200}")

    # ----------------------------------------------------------------- Ajustes
    _header(ws_adj, [("Regla", 16), ("Importe_EUR", 20), ("Explicación", 80)])
    ws_adj["E1"] = ("Partidas de conciliación explicadas (solo cuentan en reglas tipo 'conciliacion'). "
                    "Toda partida debe llevar explicación.")
    _name(wb, "A_Regla", f"Ajustes!$A$2:$A${MAX_ADJ_ROWS}")
    _name(wb, "A_Importe", f"Ajustes!$B$2:$B${MAX_ADJ_ROWS}")

    # ------------------------------------------------------------- Mapeo_hojas
    _header(ws_map, [("Marco", 9), ("Hoja", 26), ("Plantilla", 18), ("Z", 14), ("Omitir", 8)])
    ws_map["G1"] = ("Opcional. Por defecto la plantilla es el nombre de la hoja. Aquí se indica otra plantilla, "
                    "un eje Z (p. ej. hojas por clase de exposición) u Omitir = 'sí'.")
    ws_map["G1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws_map.column_dimensions["G"].width = 60

    # ------------------------------------------------------------------- Datos
    _header(ws_dat, DATOS_COLS)
    last = MAX_DATA_ROWS
    for name, col in [("D_Clave", "A"), ("D_ValorEUR", "I"), ("D_Redondeo", "K"), ("D_Fichero", "L"),
                      ("D_Hoja", "M"), ("D_Celda", "N"), ("D_Marco", "B")]:
        _name(wb, name, f"Datos!${col}$2:${col}${last}")
    ws_dat.conditional_formatting.add(f"A2:Q{last}", FormulaRule(formula=['$P2="KO"'], fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")))
    ws_dat.conditional_formatting.add(f"P2:P{last}", FormulaRule(formula=['$P2="REDONDEO"'], fill=PatternFill("solid", start_color=COLORS["REDONDEO"], end_color=COLORS["REDONDEO"])))
    ws_dat.conditional_formatting.add(f"Q2:Q{last}", FormulaRule(formula=['$Q2<>""'], fill=PatternFill("solid", start_color="FFA64D", end_color="FFA64D")))

    # Fórmulas R1C1 para la macro
    ws_f["A1"], ws_f["B1"] = "Columna", "FormulaR1C1"
    for i, (col, f) in enumerate(sorted(DATOS_FORMULAS_R1C1.items()), start=2):
        ws_f.cell(i, 1, col)
        ws_f.cell(i, 2, f.removeprefix("="))  # sin '=' para que no se evalúe aquí
    ws_f["D1"] = "Las fórmulas se guardan sin '=' inicial; la macro lo añade."

    # ------------------------------------------------------------- Operandos
    ops_cols = [("Regla", 14), ("Lado", 6), ("Coef", 6), ("Token", 38), ("Clave", 40), ("Fila_datos", 10),
                ("Encontrado", 11), ("Valor_EUR", 20), ("Redondeo_EUR", 14), ("Contribución_EUR", 20),
                ("Ruido_EUR", 14), ("Fichero", 40), ("Hoja", 16), ("Celda", 8), ("Ir_a_dato", 10),
                ("Abrir_origen", 12), ("Estado_regla", 14), ("Volver", 10)]
    _header(ws_ops, ops_cols)
    chk_row = {r.id: i for i, r in enumerate(rules, start=2)}
    first_op_row: dict[str, int] = {}
    row = 2
    for rule in rules:
        for side, expr in (("LHS", rule.lhs), ("RHS", rule.rhs)):
            for t in parse_expression(expr):
                first_op_row.setdefault(rule.id, row)
                r = row
                ws_ops.cell(r, 1, rule.id)
                ws_ops.cell(r, 2, side)
                ws_ops.cell(r, 3, t.coef)
                ws_ops.cell(r, 4, t.token)
                ws_ops.cell(r, 5, f'=IF(LEFT(D{r},8)="concept:",IFERROR(VLOOKUP(MID(D{r},9,500),K_Tabla,2,FALSE),"SIN_MAPEO"),D{r})')
                ws_ops.cell(r, 6, f"=IFERROR(MATCH(E{r},D_Clave,0),0)")
                ws_ops.cell(r, 7, f"=IF(F{r}>0,1,0)")
                ws_ops.cell(r, 8, f'=IF(F{r}>0,INDEX(D_ValorEUR,F{r}),"")')
                ws_ops.cell(r, 9, f"=IF(F{r}>0,INDEX(D_Redondeo,F{r}),0)")
                ws_ops.cell(r, 10, f"=IF(F{r}>0,C{r}*H{r},0)")
                ws_ops.cell(r, 11, f"=ABS(C{r})*I{r}")
                ws_ops.cell(r, 12, f'=IF(F{r}>0,INDEX(D_Fichero,F{r}),"")')
                ws_ops.cell(r, 13, f'=IF(F{r}>0,INDEX(D_Hoja,F{r}),"")')
                ws_ops.cell(r, 14, f'=IF(F{r}>0,INDEX(D_Celda,F{r}),"")')
                ws_ops.cell(r, 15, f'=IF(F{r}>0,HYPERLINK("#Datos!A"&(F{r}+1),"→ dato"),"falta")')
                ws_ops.cell(r, 16, f'=IF(L{r}="","",HYPERLINK(L{r}&"#\'"&M{r}&"\'!"&N{r},"abrir"))')
                ws_ops.cell(r, 17, f'=IFERROR(INDEX(C_Estado,MATCH(A{r},C_Regla,0)),"")')
                ws_ops.cell(r, 18, f'=HYPERLINK("#Checks!A{chk_row[rule.id]}","← check")')
                for c in (8, 9, 10, 11):
                    ws_ops.cell(r, c).number_format = EUR_FMT
                row += 1
    n_ops = max(row - 1, 2)
    for name, col in [("O_Regla", "A"), ("O_Lado", "B"), ("O_Clave", "E"), ("O_Encontrado", "G"),
                      ("O_Contrib", "J"), ("O_Ruido", "K"), ("O_Estado", "Q")]:
        _name(wb, name, f"Operandos!${col}$2:${col}${n_ops}")
    _status_formatting(ws_ops, f"Q2:Q{n_ops}")
    ws_ops.conditional_formatting.add(f"A2:R{n_ops}", FormulaRule(formula=["$G2=0"], fill=PatternFill("solid", start_color="EDEDED", end_color="EDEDED")))

    # ----------------------------------------------------------------- Checks
    chk_cols = [("Regla", 14), ("Descripción", 48), ("Tipo", 13), ("Origen", 11), ("Faltantes", 10),
                ("Grav_máx", 10), ("Sin_mapeo", 10), ("Faltan", 8), ("LHS_EUR", 20), ("RHS_EUR", 20),
                ("Dif_bruta_EUR", 18), ("Violación_EUR", 18), ("Ruido_EUR", 14), ("No_explicada_EUR", 18),
                ("Estado", 14), ("Rango_abs", 10), ("Rango_rel", 10), ("Rango", 8), ("Gravedad", 11),
                ("Fuente", 36), ("Cita", 60), ("LHS_expr", 40), ("RHS_expr", 40), ("Operandos", 12)]
    _header(ws_chk, chk_cols)
    for r, rule in enumerate(rules, start=2):
        src = rule.source or {}
        ws_chk.cell(r, 1, rule.id)
        ws_chk.cell(r, 2, rule.description)
        ws_chk.cell(r, 3, rule.type.value)
        ws_chk.cell(r, 4, rule.origin.value)
        ws_chk.cell(r, 5, rule.missing)
        ws_chk.cell(r, 6, rule.max_severity or "")
        ws_chk.cell(r, 7, f'=COUNTIFS(O_Regla,A{r},O_Clave,"SIN_MAPEO")')
        ws_chk.cell(r, 8, f"=COUNTIFS(O_Regla,A{r},O_Encontrado,0)")
        ne = f'OR(G{r}>0,AND(H{r}>0,E{r}<>"zero"))'  # no evaluable: importes en blanco
        ws_chk.cell(r, 9, f'=IF({ne},"",SUMIFS(O_Contrib,O_Regla,A{r},O_Lado,"LHS"))')
        ws_chk.cell(r, 10, f'=IF({ne},"",SUMIFS(O_Contrib,O_Regla,A{r},O_Lado,"RHS")'
                           f'+IF(C{r}="conciliacion",SUMIFS(A_Importe,A_Regla,A{r}),0))')
        ws_chk.cell(r, 11, f'=IF({ne},"",I{r}-J{r})')
        ws_chk.cell(r, 12, f'=IF({ne},"",IF(C{r}="menor_igual",MAX(0,K{r}),IF(C{r}="mayor_igual",MAX(0,-K{r}),ABS(K{r}))))')
        ws_chk.cell(r, 13, f"=SUMIFS(O_Ruido,O_Regla,A{r})")
        ws_chk.cell(r, 14, f'=IF({ne},"",MAX(0,L{r}-M{r}))')
        ws_chk.cell(r, 15, f'=IF({ne},"NO_EVALUABLE",IF(L{r}<=P_Tol,"OK",IF(N{r}<=P_Tol,"REDONDEO","KO")))')
        ws_chk.cell(r, 16, f'=IF(O{r}<>"KO","",IF(N{r}<P_Abs1,1,IF(N{r}<P_Abs2,2,IF(N{r}<P_Abs3,3,4))))')
        ws_chk.cell(r, 17, f'=IF(O{r}<>"KO","",IF(MAX(ABS(I{r}),ABS(J{r}))=0,0,'
                           f'IF(N{r}/MAX(ABS(I{r}),ABS(J{r}))<P_Rel1,1,IF(N{r}/MAX(ABS(I{r}),ABS(J{r}))<P_Rel2,2,'
                           f'IF(N{r}/MAX(ABS(I{r}),ABS(J{r}))<P_Rel3,3,4)))))')
        ws_chk.cell(r, 18, f'=IF(O{r}<>"KO","",MIN(MAX(P{r},Q{r}),IFERROR(MATCH(F{r},P_Niveles,0),4)))')
        ws_chk.cell(r, 19, f'=IF(O{r}<>"KO","—",INDEX(P_Niveles,R{r}))')
        ws_chk.cell(r, 20, _source_text(src))
        ws_chk.cell(r, 21, src.get("cita") or src.get("formula") or "")
        ws_chk.cell(r, 22, rule.lhs)
        ws_chk.cell(r, 23, rule.rhs)
        if rule.id in first_op_row:
            ws_chk.cell(r, 24, f'=HYPERLINK("#Operandos!A{first_op_row[rule.id]}","→ operandos")')
        for c in range(9, 15):
            ws_chk.cell(r, c).number_format = EUR_FMT
    n_chk = max(len(rules) + 1, 2)
    _name(wb, "C_Regla", f"Checks!$A$2:$A${n_chk}")
    _name(wb, "C_Origen", f"Checks!$D$2:$D${n_chk}")
    _name(wb, "C_Estado", f"Checks!$O$2:$O${n_chk}")
    _name(wb, "C_Gravedad", f"Checks!$S$2:$S${n_chk}")
    _status_formatting(ws_chk, f"O2:O{n_chk}")
    _severity_formatting(ws_chk, f"S2:S{n_chk}", sev.levels)
    ws_chk.auto_filter.ref = f"A1:X{n_chk}"

    # ----------------------------------------------------------------- Resumen
    ws_res["A1"] = "Resumen del cuadre"
    ws_res["A1"].font = Font(bold=True, size=14)
    ws_res["A2"] = f"{ex.name} · referencia {ex.reference_date}"
    ws_res["A4"], ws_res["B4"] = "Estado", "Nº reglas"
    for i, stt in enumerate(["KO", "NO_EVALUABLE", "REDONDEO", "OK"], start=5):
        ws_res.cell(i, 1, stt)
        ws_res.cell(i, 2, f'=COUNTIF(C_Estado,"{stt}")')
    _status_formatting(ws_res, "A5:A8")
    ws_res["A10"], ws_res["B10"] = "Gravedad (KO)", "Nº reglas"
    for i, lvl in enumerate(reversed(sev.levels), start=11):
        ws_res.cell(i, 1, lvl)
        ws_res.cell(i, 2, f'=COUNTIFS(C_Estado,"KO",C_Gravedad,"{lvl}")')
    _severity_formatting(ws_res, "A11:A14", sev.levels)
    ws_res["A16"], ws_res["B16"], ws_res["C16"] = "Origen", "KO", "Total"
    for i, origin in enumerate(sorted({r.origin.value for r in rules}), start=17):
        ws_res.cell(i, 1, origin)
        ws_res.cell(i, 2, f'=COUNTIFS(C_Origen,A{i},C_Estado,"KO")')
        ws_res.cell(i, 3, f"=COUNTIF(C_Origen,A{i})")
    ws_res["E4"], ws_res["F4"] = "Datos cargados", f'=COUNTIF(D_Clave,"?*")'
    ws_res["E5"], ws_res["F5"] = "Datos duplicados", f'=COUNTIF(Datos!$Q$2:$Q${MAX_DATA_ROWS},"DUPLICADO")'
    for c in ("A4", "B4", "A10", "B10", "A16", "B16", "C16", "E4", "E5"):
        ws_res[c].font = Font(bold=True)
    ws_res.column_dimensions["A"].width = 22
    ws_res.column_dimensions["E"].width = 22

    # ------------------------------------------------------------------- Carga
    _header(ws_log, [("Fecha", 18), ("Marco", 9), ("Fichero", 50), ("Hoja", 22), ("Plantilla", 18), ("Z", 12),
                     ("Filas", 7), ("Columnas", 9), ("Datos", 7), ("Avisos", 70)])

    # ------------------------------------------------------------------ Inicio
    ws_inicio["A1"] = "Cuadre Stress Test ↔ FINREP/COREP"
    ws_inicio["A1"].font = Font(bold=True, size=16)
    ws_inicio["A2"] = f"{ex.name}"
    lines = [
        "1. Revisa la hoja Parametros: escala y decimales por marco (unidad de cada fichero).",
        "2. Pulsa «Cargar ST», «Cargar FINREP» y «Cargar COREP» y elige los Excel (layout de anexo EBA: códigos de fila/columna como texto).",
        "3. Revisa la hoja Carga: filas/columnas detectadas y avisos por hoja. Si una hoja no se lee bien, ajústala en Mapeo_hojas.",
        "4. Mira Resumen y Checks. Cada check enlaza a sus Operandos; cada operando enlaza al dato y a la celda del fichero de origen.",
        "5. Las partidas de conciliación (p. ej. reexpresión CRR3) van en Ajustes, siempre con explicación.",
        "",
        "Sin macros (si el libro es .xlsx): pega los datos en la hoja Datos (columnas B–H, J, L–N) y copia las fórmulas de la fila 2.",
        "Las reglas y conceptos se mantienen en el repositorio (YAML) y este libro se regenera con: cuadre excel --ejercicio ...",
    ]
    for i, t in enumerate(lines, start=4):
        ws_inicio.cell(i, 1, t)
    ws_inicio.column_dimensions["A"].width = 140
    ws_inicio["A14"] = "Botones (los crea la macro CrearBotones):"
    ws_inicio["A14"].font = Font(italic=True, color="808080")

    # --------------------------------------------------------------- Metadatos
    meta = {
        "generado_por": f"cuadre {__version__}",
        "generado": dt.datetime.now().isoformat(timespec="seconds"),
        "ejercicio": ex.id,
        "fecha_referencia": ex.reference_date,
        "marcos": ", ".join(f"{k} {v}" for k, v in ex.frameworks.items()),
        "n_reglas": len(rules),
        "n_conceptos": len(cmap.concepts),
    }
    for f in sorted({r.source.get("_fichero_regla") for r in rules if r.source.get("_fichero_regla")}):
        meta[f"sha256 {Path(f).name}"] = sha256(Path(f))
    for i, (k, v) in enumerate(meta.items(), start=1):
        ws_meta.cell(i, 1, k).font = Font(bold=True)
        ws_meta.cell(i, 2, v)
    ws_meta.column_dimensions["A"].width = 32
    ws_meta.column_dimensions["B"].width = 80

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)

    # --------------------------------------------------------------- precarga
    if preload_sources:
        _preload(ws_dat, ws_log, ws_adj, ws_map, preload_sources, out.parent.resolve())

    wb.calculation.fullCalcOnLoad = True  # openpyxl no guarda resultados: Excel recalcula al abrir
    wb.save(out)
    return out


def _preload(ws_dat, ws_log, ws_adj, ws_map, sources_path, book_dir: Path) -> None:
    """Rellena Datos/Carga/Ajustes con los loaders Python (demo, pruebas o formatos que la
    macro no lee, como xBRL).

    Las rutas se guardan relativas al libro: Excel resuelve así los hipervínculos aunque el
    repositorio se clone en otra máquina, siempre que se mantenga la estructura de carpetas.
    """
    def rel(p: str) -> str:
        return os.path.relpath(Path(p).resolve(), book_dir)

    specs, adj_path = load_sources(sources_path)
    r = 2
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    log_r = 2
    for spec in specs:
        res = load_source(spec)
        for f in res.facts:
            vals = {2: f.ref.framework, 3: f.ref.template, 4: f.ref.row, 5: f.ref.col, 6: f.ref.z or "",
                    7: f.raw_value, 8: f.scale, 10: spec.decimals if spec.decimals is not None else "",
                    12: rel(f.source.file), 13: f.source.sheet or "", 14: f.source.cell or ""}
            for c, v in vals.items():
                ws_dat.cell(r, c, v)
            for c, fml in DATOS_FORMULAS_R1C1.items():
                ws_dat.cell(r, c, r1c1_to_a1(fml, r))
            ws_dat.cell(r, 9).number_format = EUR_FMT
            r += 1
        for c in res.coverage:
            for j, v in enumerate([now, spec.framework, rel(c.file), c.sheet, c.template, c.z,
                                   c.n_rows, c.n_cols, c.n_facts, " | ".join(c.warnings) or "(precarga Python)"], start=1):
                ws_log.cell(log_r, j, v)
            log_r += 1
        for sheet, conf in (spec.options.get("sheets") or {}).items():
            ws_map.append([spec.framework, sheet, conf.get("template", sheet), conf.get("z") or "", ""])
    for i, a in enumerate(load_adjustments(adj_path), start=2):
        ws_adj.cell(i, 1, a.rule_id)
        ws_adj.cell(i, 2, a.amount_eur).number_format = EUR_FMT
        ws_adj.cell(i, 3, a.explanation)
