"""CLI: ``cuadre run --ejercicio ... --fuentes ... --salida out/informe.xlsx``."""
from __future__ import annotations

import argparse
import sys

from .report import write_report
from .run import run


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="cuadre")
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="Ejecuta el cuadre y genera el informe Excel")
    r.add_argument("--ejercicio", required=True, help="config/exercises/*.yaml")
    r.add_argument("--fuentes", required=True, help="config/sources/*.yaml")
    r.add_argument("--salida", default="out/informe_cuadre.xlsx")
    r.add_argument("--fallar-si", choices=["nunca", "ko", "critico"], default="nunca",
                   help="Código de salida 1 si hay KO (o KO críticos): útil en CI")
    a = ap.parse_args(argv)

    res = run(a.ejercicio, a.fuentes)
    path = write_report(res, a.salida)
    counts: dict[str, int] = {}
    for x in res.results:
        counts[x.status.value] = counts.get(x.status.value, 0) + 1
    worst = [x for x in res.results if x.status.value == "KO"]
    print(f"Hechos cargados: {len(res.store)} · reglas: {len(res.results)} · " +
          " · ".join(f"{k}: {v}" for k, v in sorted(counts.items())))
    for w in res.coverage:
        for msg in w.warnings:
            print(f"  aviso carga [{w.sheet}]: {msg}", file=sys.stderr)
    print(f"Informe: {path}")
    if a.fallar_si == "ko" and worst:
        return 1
    if a.fallar_si == "critico" and any(x.severity == res.severity.levels[-1] for x in worst):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
