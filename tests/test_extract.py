import json

from openpyxl import Workbook

from cuadre.catalog import ConceptMap, load_rules
from cuadre.extract import pipeline as pl
from cuadre.extract.ingest import split_paragraphs
from cuadre.extract.inventory import build_inventory


def _write_pages(path, text):
    path.write_text(json.dumps({"doc": "TG", "file": "tg.pdf", "page": 7, "text": text,
                                "paragraphs": split_paragraphs(text, 7)}) + "\n", encoding="utf-8")


def test_paragraph_split():
    paras = split_paragraphs("Intro\n12. Banks shall report X.\n13. Another one.", 3)
    assert [p["id"] for p in paras] == ["p3.b0", "12", "13"]


def test_inventory_turns_template_formulas_into_rules(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "CSV_CAP"
    ws["C2"] = "0010"
    ws["A3"], ws["B3"], ws["C3"] = "CET1", "0010", 1
    ws["A4"], ws["B4"], ws["C4"] = "AT1", "0020", 2
    ws["A5"], ws["B5"], ws["C5"] = "Total", "0030", "=SUM(C3:C4)"
    p = tmp_path / "t.xlsx"
    wb.save(p)
    inv = build_inventory(p, "ST")
    assert len(inv["cells"]) == 3
    [rule] = inv["template_rules"]
    assert rule["lhs"] == "{ST:CSV_CAP!r0030c0010}"
    assert rule["rhs"] == "{ST:CSV_CAP!r0010c0010} + {ST:CSV_CAP!r0020c0010}"


def test_review_flow_checks_quotes_and_inventory(tmp_path):
    work = tmp_path / "work"
    (work / "agent_out/extract").mkdir(parents=True)
    (work / "agent_out/translate").mkdir(parents=True)
    _write_pages(work / "pages.jsonl", "15. The CET1 amount at the starting point shall be equal\nto the amount reported in COREP.")
    good_quote = "The CET1 amount at the starting point shall be equal to the amount reported in COREP."
    cands = [
        {"candidate_id": "c1", "doc": "TG", "page": 7, "paragraph": "15", "quote": good_quote,
         "kind": "igualdad", "operands_text": [{"side": "LHS", "text": "CET1 ST"}, {"side": "RHS", "text": "CET1 COREP"}]},
        {"candidate_id": "c2", "doc": "TG", "page": 7, "paragraph": "15", "quote": "Banks must reconcile everything with FINREP always.",
         "kind": "igualdad", "operands_text": [{"side": "LHS", "text": "x"}]},
    ]
    pl.write_jsonl(cands, work / "agent_out/extract/a.jsonl")
    fuente = {"documento": "TG", "apartado": "15", "pagina": 7, "cita": good_quote}
    drafts = [
        {"candidate_id": "c1", "rule": {"id": "ST25-CAP-1", "descripcion": "CET1", "tipo": "igualdad",
                                        "lhs": "{ST:CSV_CAP!r0010c0010}", "rhs": "{COREP:C 01.00!r0020c0010}",
                                        "origen": "normativa", "fuente": fuente}},
        {"candidate_id": "c2", "rule": {"id": "ST25-CAP-2", "descripcion": "inventada", "tipo": "igualdad",
                                        "lhs": "{ST:CSV_CAP!r0099c0010}", "rhs": "{COREP:C 01.00!r0020c0010}",
                                        "origen": "normativa", "fuente": {**fuente, "cita": "Banks must reconcile everything with FINREP always."}}},
    ]
    pl.write_jsonl(drafts, work / "agent_out/translate/a.jsonl")
    inv = {"file": "x", "framework": "ST", "template_rules": [], "cells": [
        {"ref": "ST:CSV_CAP!r0010c0010"}, {"ref": "COREP:C 01.00!r0020c0010"}]}

    review = pl.build_review(work, work / "pages.jsonl", [inv], ConceptMap())
    by_id = {r["rule"]["id"]: r for r in review}
    assert by_id["ST25-CAP-1"]["auto_ok"] is True
    bad = {c["check"]: c["ok"] for c in by_id["ST25-CAP-2"]["checks"]}
    assert bad["candidato.cita_literal"] is False
    assert bad["operandos_en_inventario"] is False

    # Decisión humana se conserva si la regla no cambia, y solo se exportan las aceptadas
    by_id["ST25-CAP-1"]["decision"] = "aceptada"
    review2 = pl.build_review(work, work / "pages.jsonl", [inv], ConceptMap(), review)
    assert {r["rule"]["id"]: r["decision"] for r in review2}["ST25-CAP-1"] == "aceptada"
    counts = pl.export_accepted(review2, tmp_path / "rules", tmp_path / "nc.yaml")
    assert counts == {"normativa": 1}
    [rule] = load_rules([tmp_path / "rules/*.yaml"])
    assert rule.id == "ST25-CAP-1" and rule.source["cita"] == good_quote


def test_extract_tasks_windows(tmp_path):
    lines = [json.dumps({"doc": "TG", "page": i, "text": f"{i}. x", "paragraphs": []}) for i in range(1, 8)]
    (tmp_path / "pages.jsonl").write_text("\n".join(lines) + "\n", encoding="utf-8")
    tasks = pl.extract_tasks(tmp_path / "pages.jsonl", window=3, overlap=1)
    assert [t["task_id"] for t in tasks] == ["extract-TG-p1-3", "extract-TG-p3-5", "extract-TG-p5-7"]
    assert all(t["prompt"].endswith("extractor.md") for t in tasks)
