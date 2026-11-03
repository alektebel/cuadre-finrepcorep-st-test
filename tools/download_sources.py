"""Descarga las fuentes de data/manifest.yaml a data/raw/<carpeta>/ y anota su SHA-256.

    python tools/download_sources.py            # todas las pendientes
    python tools/download_sources.py EBA-GL-2018-04

Prueba la URL principal y luego las alternativas. Si ninguna responde, deja la fuente
'pendiente' y lo dice; no inventa nada.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import sys
import urllib.parse
import urllib.request
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "data" / "manifest.yaml"


def _download(url: str, dest_dir: Path) -> Path:
    name = urllib.parse.unquote(Path(urllib.parse.urlparse(url).path).name) or "descarga"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / name
    req = urllib.request.Request(url, headers={"User-Agent": "cuadre-downloader/0.1"})
    with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as fh:
        fh.write(resp.read())
    return dest


def main(ids: list[str]) -> int:
    text = MANIFEST.read_text(encoding="utf-8")
    data = yaml.safe_load(text)
    failures, changed = 0, False
    for src in data["fuentes"]:
        if ids and src["id"] not in ids:
            continue
        if not ids and src.get("estado") == "descargado":
            continue
        urls = [u for u in [src.get("url"), *(src.get("url_alternativas") or [])] if u]
        if not urls:
            print(f"[{src['id']}] sin URL directa: descárgalo a mano desde {src.get('pagina')} "
                  f"a data/raw/{src['carpeta']}/")
            failures += 1
            continue
        for url in urls:
            try:
                path = _download(url, ROOT / "data" / "raw" / src["carpeta"])
            except Exception as e:  # noqa: BLE001 - se informa y se prueba la siguiente
                print(f"[{src['id']}] fallo con {url}: {e}")
                continue
            src.update(estado="descargado", fichero=str(path.relative_to(ROOT)),
                       sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                       descargado=dt.datetime.now().isoformat(timespec="seconds"), url_usada=url)
            changed = True
            print(f"[{src['id']}] OK -> {path.relative_to(ROOT)}")
            break
        else:
            failures += 1
    if not changed:
        return 1 if failures else 0
    header = "".join(line + "\n" for line in text.splitlines() if line.startswith("#"))
    MANIFEST.write_text(header + yaml.safe_dump(data, allow_unicode=True, sort_keys=False, width=100),
                        encoding="utf-8")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
