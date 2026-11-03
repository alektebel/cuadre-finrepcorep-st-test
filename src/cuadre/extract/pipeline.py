"""Pasos deterministas del flujo de extracción y contrato con los agentes.

Flujo (ver PLAN.md §8.3)::

    [1] ingest      PDF -> work/pages.jsonl
    [2] inventory   plantillas -> work/inventory_*.json (+ reglas de fórmulas internas)
    [3] tasks       -> work/tasks/extract.jsonl       ──► agentes ──► work/agent_out/extract/*.jsonl
    [4] tasks       -> work/tasks/translate.jsonl     ──► agentes ──► work/agent_out/translate/*.jsonl
    [5] review      checks deterministas + veredictos -> work/review.json
    [6] tasks       -> work/tasks/verify.jsonl        ──► agentes ──► work/agent_out/verify/*.jsonl
    [7] revisión humana (app o editando review.json) -> export -> rules/st2025/*.yaml

Los agentes son intercambiables: este módulo solo produce entradas y consume salidas
JSON Lines validadas contra ``schemas/*.json``.
"""
from __future__ import annotations

import datetime as dt
import glob
import hashlib
import json
import re
from collections import defaultdict
from importlib import resources
from pathlib import Path

import jsonschema
import yaml

from ..catalog import ConceptMap, rule_from_dict
from ..expressions import tokens
from .ingest import load_pages, normalize

PKG = resources.files("cuadre.extract")


def schema(name: str) -> dict:
    return json.loads((PKG / "schemas" / f"{name}.schema.json").read_text(encoding="utf-8"))


def prompt_path(name: str) -> str:
    return str(PKG / "prompts" / f"{name}.md")


def read_jsonl(pattern: str | Path) -> list[dict]:
    out = []
    for f in sorted(glob.glob(str(pattern))):
        with open(f, encoding="utf-8") as fh:
            for i, line in enumerate(fh, 1):
                if line.strip():
                    d = json.loads(line)
                    d["_file"] = f"{f}:{i}"
                    out.append(d)
    return out


def write_jsonl(rows, path: str | Path) -> int:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    n = 0
    with open(path, "w", encoding="utf-8") as fh:
        for r in rows:
            fh.write(json.dumps(r, ensure_ascii=False) + "\n")
            n += 1
    return n


def _clean(d: dict) -> dict:
    return {k: v for k, v in d.items() if not k.startswith("_")}


# --------------------------------------------------------------------------- tareas
MENTION_RE = re.compile(r"FINREP|COREP|reconcil|consistent with|in line with .{0,40}report", re.I)


def extract_tasks(pages_path: str | Path, window: int = 3, overlap: int = 1, work: Path = Path("work"),
                  docs: list[str] | None = None, only_mentions: bool = False) -> list[dict]:
    """Ventanas de páginas para el extractor.

    ``only_mentions`` descarta ventanas sin ninguna mención a FINREP/COREP/reconciliación:
    abarata la pasada, a costa de perder cuadres que el texto describa sin nombrarlos.
    """
    by_doc = defaultdict(list)
    for (doc, _), p in sorted(load_pages(pages_path).items()):
        if docs and doc not in docs:
            continue
        by_doc[doc].append(p)
    tasks = []
    for doc, pages in by_doc.items():
        step = max(1, window - overlap)
        for start in range(0, len(pages), step):
            chunk = pages[start:start + window]
            tid = f"extract-{doc}-p{chunk[0]['page']}-{chunk[-1]['page']}"
            if only_mentions and not any(MENTION_RE.search(p["text"]) for p in chunk):
                if start + window >= len(pages):
                    break
                continue
            tasks.append({
                "task_id": tid, "stage": "extract",
                "prompt": prompt_path("extractor"), "output_schema": "candidate",
                "output": str(work / "agent_out/extract" / f"{tid}.jsonl"),
                "input": {"doc": doc, "pages": [{"page": p["page"], "paragraphs": p["paragraphs"]} for p in chunk]},
            })
            if start + window >= len(pages):
                break
    return tasks


def translate_tasks(candidates: list[dict], context_files: list[str], work: Path = Path("work")) -> list[dict]:
    return [{
        "task_id": f"translate-{c['candidate_id']}", "stage": "translate",
        "prompt": prompt_path("translator"), "output_schema": "rule_draft",
        "output": str(work / "agent_out/translate" / f"{c['candidate_id']}.jsonl"),
        "input": {"candidate": _clean(c)}, "context_files": context_files,
    } for c in candidates]


def verify_tasks(review: list[dict], labels: dict[str, str], work: Path = Path("work")) -> list[dict]:
    out = []
    for item in review:
        r = item["rule"]
        refs = {}
        for t in tokens(r["lhs"]) + tokens(r["rhs"]):
            refs[t] = labels.get(t) or labels.get(item.get("resolved", {}).get(t, ""), None)
        out.append({
            "task_id": f"verify-{r['id']}", "stage": "verify",
            "prompt": prompt_path("verifier"), "output_schema": "verdict",
            "output": str(work / "agent_out/verify" / f"{r['id']}.jsonl"),
            "input": {"rule": {k: r[k] for k in ("id", "tipo", "lhs", "rhs")},
                      "cita": (r.get("fuente") or {}).get("cita") or (r.get("fuente") or {}).get("formula"),
                      "etiquetas": refs},
        })
    return out


# --------------------------------------------------------------- checks deterministas
def _check(name: str, ok: bool, detail: str = "") -> dict:
    return {"check": name, "ok": bool(ok), "detail": detail}


def quote_in_pages(quote: str, doc: str, page: int, pages: dict) -> bool:
    q = normalize(quote)
    for p in (page, page + 1, page - 1):  # citas que cruzan salto de página
        texts = [pages.get((doc, p), {}).get("text", "")]
        if p == page:
            texts.append(texts[0] + " " + pages.get((doc, page + 1), {}).get("text", ""))
        if any(q in normalize(t) for t in texts):
            return True
    return False


def check_candidate(c: dict, pages: dict) -> list[dict]:
    checks = []
    try:
        jsonschema.validate(_clean(c), schema("candidate"))
        checks.append(_check("esquema", True))
    except jsonschema.ValidationError as e:
        return [_check("esquema", False, e.message)]
    checks.append(_check("pagina_existe", (c["doc"], c["page"]) in pages, f"{c['doc']} p.{c['page']}"))
    checks.append(_check("cita_literal", quote_in_pages(c["quote"], c["doc"], c["page"], pages),
                         "la cita debe aparecer literalmente en la página indicada"))
    return checks


def check_draft(d: dict, pages: dict, refs: set[str], cmap: ConceptMap) -> tuple[list[dict], dict]:
    checks, resolved = [], {}
    try:
        jsonschema.validate(_clean(d), schema("rule_draft"))
        checks.append(_check("esquema", True))
    except jsonschema.ValidationError as e:
        return [_check("esquema", False, e.message)], resolved
    try:
        rule_from_dict(d["rule"])
        checks.append(_check("dsl_valido", True))
    except Exception as e:  # noqa: BLE001 - se informa cualquier fallo de parseo
        return checks + [_check("dsl_valido", False, str(e))], resolved

    new = d.get("new_concepts") or {}
    missing = []
    for t in tokens(d["rule"]["lhs"]) + tokens(d["rule"]["rhs"]):
        if t.startswith("concept:"):
            name = t.removeprefix("concept:")
            ref = str(cmap.concepts[name].ref) if name in cmap.concepts else (new.get(name) or {}).get("ref")
        else:
            ref = t
        if ref is None:
            missing.append(f"{t} (concepto inexistente)")
            continue
        resolved[t] = ref
        if refs and ref not in refs:
            missing.append(f"{t} -> {ref} (celda fuera de inventario)")
    checks.append(_check("operandos_en_inventario", not missing, "; ".join(missing)))
    checks.append(_check("sin_operandos_sin_resolver", not d.get("unresolved"),
                         ", ".join(d.get("unresolved") or [])))
    src = d["rule"].get("fuente") or {}
    if d["rule"]["origen"] == "normativa":
        ok = bool(src.get("cita")) and quote_in_pages(src["cita"], src.get("documento", ""), int(src.get("pagina") or 0), pages)
        checks.append(_check("cita_normativa_literal", ok, f"{src.get('documento')} p.{src.get('pagina')}"))
    return checks, resolved


# ------------------------------------------------------------------------- revisión
def _rule_hash(rule: dict) -> str:
    return hashlib.sha256(json.dumps(rule, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:16]


def build_review(work: Path, pages_path: Path | None, inventories: list[dict], cmap: ConceptMap,
                 previous: list[dict] | None = None) -> list[dict]:
    pages = load_pages(pages_path) if pages_path and pages_path.exists() else {}
    refs = {c["ref"] for inv in inventories for c in inv["cells"]}
    candidates = {c["candidate_id"]: c for c in read_jsonl(work / "agent_out/extract/*.jsonl")}
    cand_checks = {cid: check_candidate(c, pages) for cid, c in candidates.items()}
    verdicts = {v["rule_id"]: v for v in read_jsonl(work / "agent_out/verify/*.jsonl")}
    prev = {p["rule"]["id"]: p for p in previous or []}

    drafts = read_jsonl(work / "agent_out/translate/*.jsonl") + read_jsonl(work / "agent_out/propose/*.jsonl")
    for inv in inventories:  # reglas de fórmulas de plantilla: no pasan por agentes
        drafts += [{"candidate_id": f"TPL:{r['_celda']}", "rule": {k: v for k, v in r.items() if not k.startswith("_")},
                    "_file": inv["file"]} for r in inv.get("template_rules", [])]

    out = []
    for d in drafts:
        checks, resolved = check_draft(d, pages, refs, cmap)
        cid = d.get("candidate_id")
        if cid in cand_checks:
            checks = [dict(c, check=f"candidato.{c['check']}") for c in cand_checks[cid]] + checks
        rule = d.get("rule") or {}
        h = _rule_hash(rule)
        p = prev.get(rule.get("id"))
        keep = p is not None and p.get("rule_hash") == h
        out.append({
            "rule": rule, "rule_hash": h, "candidate_id": cid, "origin_file": d.get("_file"),
            "new_concepts": d.get("new_concepts") or {}, "resolved": resolved,
            "checks": checks, "auto_ok": all(c["ok"] for c in checks),
            "verdict": verdicts.get(rule.get("id")),
            "decision": p["decision"] if keep else "pendiente",
            "comment": p.get("comment", "") if keep else "",
            "reviewed_at": p.get("reviewed_at") if keep else None,
        })
    return out


def export_accepted(review: list[dict], rules_dir: Path, concepts_out: Path) -> dict[str, int]:
    by_origin = defaultdict(list)
    new_concepts = {}
    for item in review:
        if item["decision"] != "aceptada":
            continue
        if not item["auto_ok"]:
            raise ValueError(f"{item['rule']['id']}: aceptada pero con checks automáticos fallidos")
        by_origin[item["rule"]["origen"]].append(item["rule"])
        new_concepts.update(item.get("new_concepts") or {})
    rules_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.date.today().isoformat()
    for origin, rules in by_origin.items():
        name = {"normativa": "normativa", "plantilla": "plantilla", "propuesta": "propuestas"}[origin]
        (rules_dir / f"{name}.yaml").write_text(
            f"# Generado por cuadre-extract export el {stamp}. Editar en work/review.json y reexportar.\n"
            + yaml.safe_dump({"reglas": rules}, allow_unicode=True, sort_keys=False, width=120),
            encoding="utf-8")
    if new_concepts:
        concepts_out.parent.mkdir(parents=True, exist_ok=True)
        concepts_out.write_text(
            "# Conceptos nuevos de reglas aceptadas: revisar y fusionar en concepts/*.yaml\n"
            + yaml.safe_dump({"conceptos": {k: {**v, "verificado": False} for k, v in new_concepts.items()}},
                             allow_unicode=True, sort_keys=True), encoding="utf-8")
    return {k: len(v) for k, v in by_origin.items()}


# ------------------------------------------------------------------- propuestas
_WORD = re.compile(r"[a-záéíóúñ0-9]+")
_STOP = {"of", "the", "and", "for", "to", "in", "de", "la", "el", "y", "los", "las", "del", "total"}


def _words(s: str | None) -> set[str]:
    return {w for w in _WORD.findall((s or "").lower()) if w not in _STOP and len(w) > 2}


def _lbl(cell: dict) -> str:
    return " / ".join(x for x in (cell.get("row_label"), cell.get("col_label")) if x)


def propose_by_labels(st_inv: dict, reg_invs: list[dict], threshold: float = 0.6) -> list[dict]:
    """Empareja celdas ST con celdas FINREP/COREP por similitud (Jaccard) de etiquetas.

    Es una heurística ruidosa: sirve como punto de partida para el agente proponedor y
    para revisión humana, nunca como regla aceptada automáticamente.
    """
    reg = [(c, _words(f"{c['row_label']} {c['col_label']}")) for inv in reg_invs for c in inv["cells"]]
    out, n = [], 0
    for c in st_inv["cells"]:
        w = _words(f"{c['row_label']} {c['col_label']}")
        if not w:
            continue
        best, score = None, 0.0
        for rc, rw in reg:
            if not rw:
                continue
            s = len(w & rw) / len(w | rw)
            if s > score:
                best, score = rc, s
        if best and score >= threshold:
            n += 1
            out.append({
                "candidate_id": f"PROP-LBL-{n}",
                "rule": {
                    "id": f"PROP-LBL-{n}",
                    "descripcion": f"(Propuesta por etiqueta, similitud {score:.2f}) '{_lbl(c)}' ≈ '{_lbl(best)}'",
                    "tipo": "igualdad", "lhs": "{" + c["ref"] + "}", "rhs": "{" + best["ref"] + "}",
                    "origen": "propuesta", "gravedad_maxima": "Leve",
                },
            })
    return out
