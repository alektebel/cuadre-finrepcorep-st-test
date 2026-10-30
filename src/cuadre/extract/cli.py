"""CLI del flujo de extracción: ``cuadre-extract <paso> ...``.

Ejemplo completo::

    cuadre-extract ingest --pdf data/raw/eba_st2025/template_guidance.pdf --doc TG2025
    cuadre-extract inventory --xlsx data/raw/eba_st2025/templates.xlsx --marco ST --out work/inventory_st.json
    cuadre-extract tasks extract                 # -> work/tasks/extract.jsonl (para tus agentes)
    cuadre-extract tasks translate               # tras volcar candidatos en work/agent_out/extract/
    cuadre-extract review                        # checks deterministas -> work/review.json
    cuadre-extract tasks verify                  # -> work/tasks/verify.jsonl
    cuadre-extract review                        # incorpora veredictos
    #   ... revisión humana en la app (pestaña "Revisión de reglas") ...
    cuadre-extract export --rules-dir rules/st2025
"""
from __future__ import annotations

import argparse
import glob
import json
from pathlib import Path

from ..catalog import ConceptMap
from . import pipeline as pl
from .ingest import ingest_pdf
from .inventory import build_inventory, save_inventory


def _inventories(work: Path) -> list[dict]:
    return [json.loads(Path(f).read_text(encoding="utf-8")) for f in sorted(glob.glob(str(work / "inventory_*.json")))]


def _labels(invs: list[dict]) -> dict[str, str]:
    return {c["ref"]: pl._lbl(c) for inv in invs for c in inv["cells"]}


def load_review(work: Path) -> list[dict]:
    p = work / "review.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else []


def save_review(work: Path, review: list[dict]) -> None:
    (work / "review.json").write_text(json.dumps(review, ensure_ascii=False, indent=1), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="cuadre-extract")
    ap.add_argument("--work", default="work", help="carpeta de trabajo")
    ap.add_argument("--conceptos", action="append",
                    help="mapa de conceptos (repetible); por defecto concepts/*.yaml")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("ingest", help="[1] PDF -> work/pages.jsonl")
    s.add_argument("--pdf", required=True)
    s.add_argument("--doc", help="id corto del documento (se usa en las citas)")

    s = sub.add_parser("inventory", help="[2] plantillas Excel -> inventario de celdas")
    s.add_argument("--xlsx", required=True)
    s.add_argument("--marco", required=True, help="ST | FINREP | COREP")
    s.add_argument("--out", required=True, help="work/inventory_<algo>.json")

    s = sub.add_parser("tasks", help="genera tareas JSONL para los agentes")
    s.add_argument("stage", choices=["extract", "translate", "verify"])
    s.add_argument("--window", type=int, default=3)

    sub.add_parser("propose", help="[P] propuestas deterministas por similitud de etiquetas")
    sub.add_parser("review", help="[5] checks deterministas + veredictos -> work/review.json")

    s = sub.add_parser("export", help="[7] reglas aceptadas -> rules/")
    s.add_argument("--rules-dir", default="rules/st2025")

    sub.add_parser("status", help="resumen del estado de la revisión")
    a = ap.parse_args(argv)
    work = Path(a.work)
    work.mkdir(parents=True, exist_ok=True)
    a.conceptos = a.conceptos or sorted(glob.glob("concepts/*.yaml"))

    if a.cmd == "ingest":
        n = ingest_pdf(a.pdf, work / "pages.jsonl", a.doc)
        print(f"{n} páginas añadidas a {work / 'pages.jsonl'}")
    elif a.cmd == "inventory":
        inv = build_inventory(a.xlsx, a.marco)
        save_inventory(inv, a.out)
        print(f"{len(inv['cells'])} celdas, {len(inv['template_rules'])} reglas de fórmula, "
              f"{len(inv['issues'])} incidencias -> {a.out}")
    elif a.cmd == "tasks":
        if a.stage == "extract":
            tasks = pl.extract_tasks(work / "pages.jsonl", window=a.window, work=work)
        elif a.stage == "translate":
            pages = pl.load_pages(work / "pages.jsonl") if (work / "pages.jsonl").exists() else {}
            ok = [c for c in pl.read_jsonl(work / "agent_out/extract/*.jsonl")
                  if all(x["ok"] for x in pl.check_candidate(c, pages))]
            ctx = sorted(glob.glob(str(work / "inventory_*.json"))) + list(a.conceptos)
            tasks = pl.translate_tasks(ok, ctx, work=work)
        else:
            tasks = pl.verify_tasks(load_review(work), _labels(_inventories(work)), work=work)
        n = pl.write_jsonl(tasks, work / "tasks" / f"{a.stage}.jsonl")
        print(f"{n} tareas -> {work / 'tasks' / f'{a.stage}.jsonl'}")
    elif a.cmd == "propose":
        invs = _inventories(work)
        st = [i for i in invs if i["framework"] == "ST"]
        reg = [i for i in invs if i["framework"] != "ST"]
        props = [p for s in st for p in pl.propose_by_labels(s, reg)]
        out = work / "agent_out/propose/labels.jsonl"
        n = pl.write_jsonl(props, out)
        print(f"{n} propuestas por etiqueta -> {out}")
    elif a.cmd == "review":
        cmap = ConceptMap.load([Path(p) for p in a.conceptos])
        review = pl.build_review(work, work / "pages.jsonl", _inventories(work), cmap, load_review(work))
        save_review(work, review)
        ok = sum(r["auto_ok"] for r in review)
        print(f"{len(review)} reglas en revisión ({ok} pasan checks automáticos) -> {work / 'review.json'}")
    elif a.cmd == "export":
        counts = pl.export_accepted(load_review(work), Path(a.rules_dir), work / "export" / "conceptos_nuevos.yaml")
        print(f"Exportadas: {counts or 'ninguna (no hay reglas aceptadas)'}")
    elif a.cmd == "status":
        review = load_review(work)
        from collections import Counter
        print(Counter((r["rule"].get("origen"), r["decision"], r["auto_ok"]) for r in review))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
