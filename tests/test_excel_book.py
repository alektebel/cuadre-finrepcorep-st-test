import re
import shutil
import subprocess

import pytest
from openpyxl import load_workbook

from cuadre.excel_book import DATOS_FORMULAS_R1C1, build_workbook, r1c1_to_a1
from cuadre.excel_book.generator import DATOS_COLS
from cuadre.model import Status

EXERCISE = "config/exercises/synthetic.yaml"


def test_r1c1_to_a1():
    assert r1c1_to_a1("=RC7*RC8", 5) == "=G5*H5"
    assert r1c1_to_a1('=IF(RC10="",0,RC12)', 12) == '=IF(J12="",0,L12)'


def test_structure_and_names(root, tmp_path):
    out = build_workbook(root / EXERCISE, tmp_path / "libro.xlsx")
    wb = load_workbook(out)
    assert {"Inicio", "Resumen", "Checks", "Operandos", "Datos", "Carga", "Mapeo_hojas", "Conceptos",
            "Ajustes", "Parametros", "Metadatos", "_Formulas"} <= set(wb.sheetnames)
    assert wb["_Formulas"].sheet_state == "hidden"
    names = set(wb.defined_names.keys())
    assert {"D_Clave", "D_ValorEUR", "O_Regla", "O_Contrib", "C_Estado", "P_Tol", "P_Niveles", "K_Tabla"} <= names
    f = {r[0]: r[1] for r in wb["_Formulas"].iter_rows(min_row=2, values_only=True) if r[0]}
    assert f == {k: v.removeprefix("=") for k, v in DATOS_FORMULAS_R1C1.items()}
    assert wb["Datos"].max_row == 1  # sin precarga no hay datos


def _vba(root) -> str:
    return (root / "src/cuadre/excel_book/vba/Cuadre.bas").read_bytes().decode("ascii")


def test_vba_is_consistent_with_generator(root, tmp_path):
    src = _vba(root)
    assert "\r\n" in src and "\n" not in src.replace("\r\n", "")
    wb = load_workbook(build_workbook(root / EXERCISE, tmp_path / "libro.xlsx"))
    sheets = re.findall(r'Private Const HOJA_\w+ As String = "([^"]+)"', src)
    assert sheets and set(sheets) <= set(wb.sheetnames)
    assert int(re.search(r"NCOLS_DATOS As Long = (\d+)", src).group(1)) == len(DATOS_COLS)
    macros = re.search(r'macros = Array\(([^)]*)\)', src).group(1)
    for m in re.findall(r'"(\w+)"', macros):
        assert re.search(rf"^Public Sub {m}\(", src, re.M), m
    # nombres definidos que usa la macro
    for name in re.findall(r'Names\("(\w+)"\)', src):
        assert name in wb.defined_names, name
    # ParametroMarco lee desde la fila 13 de Parametros, donde el generador pone los marcos
    assert wb["Parametros"]["A13"].value == "ST"


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice no disponible")
def test_formulas_match_python_engine(root, tmp_path, synthetic_sources, synthetic_run):
    """El libro recalculado por LibreOffice debe dar el mismo resultado que el motor Python."""
    src_xlsx = build_workbook(root / EXERCISE, tmp_path / "libro.xlsx", synthetic_sources)
    outdir = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(outdir), str(src_xlsx)],
                   check=True, capture_output=True, timeout=180)
    ws = load_workbook(outdir / "libro.xlsx", data_only=True)["Checks"]
    head = [c.value for c in ws[1]]
    got = {row[0]: dict(zip(head, row)) for row in ws.iter_rows(min_row=2, values_only=True) if row[0]}

    run, _ = synthetic_run
    assert set(got) == {r.rule.id for r in run.results}
    for r in run.results:
        g = got[r.rule.id]
        assert g["Estado"] == r.status.value, r.rule.id
        assert g["Gravedad"] == r.severity, r.rule.id
        if r.status != Status.NO_EVALUABLE:
            assert g["LHS_EUR"] == pytest.approx(r.lhs, abs=1e-6), r.rule.id
            assert g["RHS_EUR"] == pytest.approx(r.rhs, abs=1e-6), r.rule.id
            assert g["No_explicada_EUR"] == pytest.approx(r.diff_unexplained, abs=1e-6), r.rule.id
        else:
            assert g["LHS_EUR"] in (None, "")
