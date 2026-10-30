from openpyxl import load_workbook

from cuadre.model import Status
from cuadre.report import write_report


def test_synthetic_breaks_are_exactly_the_expected_ones(synthetic_run):
    run, expected = synthetic_run
    got = {r.rule.id: {"estado": r.status.value, "gravedad": r.severity}
           for r in run.results if r.status != Status.OK}
    assert got == expected


def test_every_operand_is_traceable(synthetic_run):
    run, _ = synthetic_run
    for r in run.results:
        for o in r.operands:
            if o.fact is not None:
                assert o.fact.source.file and (o.fact.source.cell or o.fact.source.detail)


def test_loader_coverage_has_no_warnings(synthetic_run):
    run, _ = synthetic_run
    assert [c for c in run.coverage if c.warnings] == []
    assert all(c.n_facts > 0 for c in run.coverage)


def test_excel_report_links(synthetic_run, tmp_path):
    run, _ = synthetic_run
    out = write_report(run, tmp_path / "informe.xlsx")
    wb = load_workbook(out)
    assert {"Resumen", "Resultados", "Detalle", "No_evaluables", "Metadatos"} <= set(wb.sheetnames)
    det = wb["Detalle"]
    header = [c.value for c in det[1]]
    col = header.index("abrir_origen") + 1
    links = [det.cell(i, col).hyperlink.target for i in range(2, det.max_row + 1) if det.cell(i, col).hyperlink]
    assert links and all("#'" in link and "!" in link for link in links)
    assert wb["Resultados"].cell(2, 1).hyperlink.location.startswith("'Detalle'!A")
