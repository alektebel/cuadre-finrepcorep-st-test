"""Plantillas oficiales ST (layout RowNum): extracción del mapeo COREP, carga y cuadre."""
import datetime as dt
import re

import yaml
from openpyxl import Workbook

from cuadre.catalog import ConceptMap, rule_from_dict
from cuadre.engine import evaluate
from cuadre.extract.st_mapping import extract
from cuadre.loaders import SourceSpec, load_source
from cuadre.model import Status
from cuadre.severity import SeverityConfig
from cuadre.st_template import parse_full_cells
from cuadre.store import FactStore


def _mini_template(path, values: bool):
    """Imita CSV_CAP (por periodos) y CSV_CR_REA_STA_DET (filas Actual/proyección, mapeo fila × columna)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "CSV_CAP"
    for j, n in enumerate(range(1, 4)):
        ws.cell(3, 12 + j, n)
    ws["L4"], ws["M4"], ws["N4"] = "Actual", "Baseline Scenario", "Adverse Scenario"
    ws["B6"], ws["K6"], ws["L6"], ws["M6"], ws["N6"], ws["S6"] = "RowNum", "(mln EUR,%)", dt.datetime(2026, 12, 31), 2027, 2027, "COREP CODE / COMPUTATION"
    rows = [
        (1, "OWN FUNDS", "C 01.00 (r0010,c0010)", 5251.0),
        (2, "CET1", "C 01.00 (r0020,c0010)", 4650.0),
        (3, "Instruments", "C 01.00 (r0030,c0010) + C 01.00 (r0040,c0010)", 2000.0),
        (4, "Ratio", "C 03.00 (r0010,c0010) %", 0.15),
        (5, "Internal", "[A.1-A.1.13]", 1.0),
    ]
    for i, (rn, lab, m, v) in enumerate(rows, start=7):
        ws.cell(i, 2, rn)
        ws.cell(i, 11, lab)
        ws.cell(i, 19, m)
        if values:
            ws.cell(i, 12, v)
            ws.cell(i, 13, v * 0.9)  # proyección: no debe generar regla ni importar al cuadre
    ws2 = wb.create_sheet("CSV_CR_REA_STA_DET")
    ws2["M4"], ws2["N4"] = "1", "2"
    ws2["M8"], ws2["N8"] = " C 09.01, c0075", "C 09.01, c0055+c0061"
    hdr = ["RowNum", "Asset class identifier", "Unique row identifier", "Geographical breakdown", "Scenario", "Year",
           "Asset classes", "Exposure value", "Adjustments", "COREP CODE / COMPUTATION"]
    for j, h in enumerate(hdr, start=2):  # B..K ; ojo: valores en M, N
        ws2.cell(9, j, h)
    ws2.cell(9, 13, "Exposure value")
    ws2.cell(9, 14, "Adjustments")
    ws2.cell(9, 16, "COREP CODE / COMPUTATION")
    data = [
        (1, "Total", "Actual", " C 09.01, r0010", 100.0, 5.0),
        (2, "Spain", "Actual", "C 09.02, r0020 or C 09.01, r0020", 60.0, 3.0),
        (3, "Total", "Baseline", " projection(C 09.01, r0010)", 90.0, 4.0),
    ]
    for i, (rn, geo, scen, m, v1, v2) in enumerate(data, start=10):
        ws2.cell(i, 2, str(rn))
        ws2.cell(i, 5, geo)
        ws2.cell(i, 6, scen)
        ws2.cell(i, 8, "Corporates")
        ws2.cell(i, 16, m)
        if values:
            ws2.cell(i, 13, v1)
            ws2.cell(i, 14, v2)
    wb.save(path)


def test_parse_full_cells():
    assert parse_full_cells("C 01.00 (r0030,c0010) + C 01.00 (r0040 ; c0010)") == [
        (1.0, "C 01.00", "0030", "0010"), (1.0, "C 01.00", "0040", "0010")]
    assert parse_full_cells("C 04.00 (r0096,c0010) * 2.5 + C 04.00 (r0504,c0010) * 2.5")[0][0] == 2.5
    assert parse_full_cells("C 01.00 (r0010,c0010) \n[A.4.1+A.4.2]") == [(1.0, "C 01.00", "0010", "0010")]
    assert parse_full_cells("C 03.00 (r0010,c0010) %") is None
    assert parse_full_cells("[A.1-A.1.13]") is None


def test_extract_mapping_rules(tmp_path):
    p = tmp_path / "tpl.xlsx"
    _mini_template(p, values=False)
    res = extract(p, prefix="T")
    rules = {r["id"]: r for r in res["rules"] if "-MAP-" in r["id"]}
    # CSV_CAP: solo columna Actual (c0001); % y cálculo interno no generan regla
    assert rules["T-MAP-CSV_CAP-r0003c0001"]["rhs"] == "{COREP:C 01.00!r0030c0010} + {COREP:C 01.00!r0040c0010}"
    assert not any(k.startswith("T-MAP-CSV_CAP-r0004") or k.startswith("T-MAP-CSV_CAP-r0005") for k in rules)
    assert not any(re.search(r"CSV_CAP-r\d+c0002", k) for k in rules)
    # Riesgo de crédito: fila × columna, z geográfico, sumas de columnas, sin proyecciones
    assert rules["T-MAP-CSV_CR_REA_STA_DET-r0001c0001"]["rhs"] == "{COREP:C 09.01!r0010c0075@GEO=Total}"
    assert rules["T-MAP-CSV_CR_REA_STA_DET-r0001c0002"]["rhs"] == \
        "{COREP:C 09.01!r0010c0055@GEO=Total} + {COREP:C 09.01!r0010c0061@GEO=Total}"
    # 'C 09.02 or C 09.01': solo C 09.01 tiene columnas en la cabecera -> sin sufijo de alternativa
    assert rules["T-MAP-CSV_CR_REA_STA_DET-r0002c0001"]["rhs"] == "{COREP:C 09.01!r0020c0075@GEO=Spain}"
    assert not any("r0003" in k for k in rules if "STA_DET" in k)
    assert res["issues"]["proyeccion"] == 1 and res["issues"]["porcentaje"] == 1
    for r in rules.values():
        rule_from_dict(r)  # DSL válido
        assert r["fuente"]["cita"] and r["origen"] == "plantilla"


def test_st_eba_loader_and_reconciliation(tmp_path):
    tpl, filled = tmp_path / "tpl.xlsx", tmp_path / "filled.xlsx"
    _mini_template(tpl, values=False)
    _mini_template(filled, values=True)
    rules = [rule_from_dict(r) for r in extract(tpl, prefix="T")["rules"] if "-MAP-" in r["id"]]

    store = FactStore()
    st = load_source(SourceSpec("ST", "st_eba", filled, scale=1e6, decimals=2))
    store.extend(st.facts)
    cov = {c.sheet: c for c in st.coverage}
    assert cov["CSV_CAP"].n_facts == 10 and cov["CSV_CR_REA_STA_DET"].n_facts == 6

    wb = Workbook()
    ws = wb.active
    ws.title = "C 01.00"
    ws["C2"] = "0010"
    for i, (code, v) in enumerate([("0010", 5251e6), ("0020", 4650e6 - 2e6), ("0030", 1500e6), ("0040", 500e6)], start=3):
        ws.cell(i, 2, code)
        ws.cell(i, 3, v)
    ws = wb.create_sheet("C 09.01 Total")
    ws["C2"], ws["D2"], ws["E2"] = "0055", "0061", "0075"
    ws["B3"], ws["C3"], ws["D3"], ws["E3"] = "0010", 4e6, 1e6, 100e6
    corep = tmp_path / "corep.xlsx"
    wb.save(corep)
    store.extend(load_source(SourceSpec("COREP", "excel_annex", corep, decimals=2, options={
        "sheets": {"C 01.00": {"template": "C 01.00"}, "C 09.01 Total": {"template": "C 09.01", "z": "GEO=Total"}}})).facts)

    res = {r.rule.id: r for r in evaluate(rules, store, ConceptMap(), SeverityConfig())}
    assert res["T-MAP-CSV_CAP-r0001c0001"].status == Status.OK
    assert res["T-MAP-CSV_CAP-r0003c0001"].status == Status.OK
    ko = res["T-MAP-CSV_CAP-r0002c0001"]
    assert ko.status == Status.KO and ko.severity == "Crítico" and abs(ko.diff_raw - 2e6) < 1e-6
    assert res["T-MAP-CSV_CR_REA_STA_DET-r0001c0001"].status == Status.OK
    assert res["T-MAP-CSV_CR_REA_STA_DET-r0001c0002"].status == Status.OK
    assert res["T-MAP-CSV_CR_REA_STA_DET-r0002c0001"].status == Status.NO_EVALUABLE  # sin hoja COREP de España


def test_repo_rules_st2027_load(root):
    """Las reglas generadas desde la plantilla oficial 2027 cargan y son todas de mapeo con cita."""
    f = root / "rules/st2027/plantilla_mapeo_corep.yaml"
    if not f.exists():
        return
    data = yaml.safe_load(f.read_text(encoding="utf-8"))["reglas"]
    rules = [rule_from_dict(d) for d in data]
    assert len(rules) == len({r.id for r in rules}) > 2000
    assert all(r.source.get("cita") and r.lhs.startswith("{ST:") and "COREP:" in r.rhs for r in rules)
