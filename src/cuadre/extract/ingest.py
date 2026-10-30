"""Paso 1: PDF (guía / nota metodológica) -> texto por página y párrafo numerado.

Los documentos EBA numeran los párrafos (``123. Banks shall...``). Se conserva ese
número porque es lo que se cita en la regla (``fuente.apartado``). Si un bloque no
lleva número, se identifica como ``p<página>.b<n>``.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from pypdf import PdfReader

_PARA_RE = re.compile(r"^\s*(\d{1,4})\.\s+(?=\S)", re.M)


def normalize(text: str) -> str:
    """Normalización usada para comparar citas literalmente."""
    text = text.replace("­", "").replace("’", "'").replace("‘", "'")
    text = text.replace("“", '"').replace("”", '"').replace("–", "-").replace("—", "-")
    text = re.sub(r"-\s*\n\s*", "", text)  # guiones de partición de línea
    return re.sub(r"\s+", " ", text).strip().lower()


def split_paragraphs(page_text: str, page: int) -> list[dict]:
    marks = list(_PARA_RE.finditer(page_text))
    paras = []
    if not marks:
        blocks = [b for b in re.split(r"\n\s*\n", page_text) if b.strip()]
        return [{"id": f"p{page}.b{i}", "text": b.strip()} for i, b in enumerate(blocks, 1)]
    if marks[0].start() > 0 and page_text[: marks[0].start()].strip():
        paras.append({"id": f"p{page}.b0", "text": page_text[: marks[0].start()].strip()})
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(page_text)
        paras.append({"id": m.group(1), "text": page_text[m.end():end].strip()})
    return paras


def ingest_pdf(pdf: str | Path, out_jsonl: str | Path, doc_id: str | None = None) -> int:
    pdf = Path(pdf)
    doc_id = doc_id or pdf.stem
    reader = PdfReader(str(pdf))
    out = Path(out_jsonl)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "a", encoding="utf-8") as fh:
        for i, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            fh.write(json.dumps({"doc": doc_id, "file": str(pdf), "page": i, "text": text,
                                 "paragraphs": split_paragraphs(text, i)}, ensure_ascii=False) + "\n")
    return len(reader.pages)


def load_pages(path: str | Path) -> dict[tuple[str, int], dict]:
    pages = {}
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            p = json.loads(line)
            pages[(p["doc"], p["page"])] = p
    return pages
