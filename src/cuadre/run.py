"""Orquestación de una ejecución: ejercicio + fuentes -> resultados trazables."""
from __future__ import annotations

import datetime as dt
import hashlib
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from . import __version__
from .catalog import Adjustment, ConceptMap, load_adjustments, load_rules
from .engine import evaluate
from .loaders import CoverageEntry, SourceSpec, load_source
from .model import Rule, RuleResult
from .severity import SeverityConfig
from .store import FactStore


def project_root(start: Path) -> Path:
    for p in [start, *start.parents]:
        if (p / "pyproject.toml").exists():
            return p
    return start


def _resolve(p: str | Path, base: Path) -> Path:
    p = Path(p)
    return p if p.is_absolute() else base / p


@dataclass
class Exercise:
    id: str
    name: str
    reference_date: str
    frameworks: dict[str, str]
    concept_files: list[Path]
    rule_patterns: list[Path]
    severity_file: Path | None
    file: Path

    @classmethod
    def load(cls, path: str | Path) -> "Exercise":
        path = Path(path).resolve()
        root = project_root(path.parent)
        d = yaml.safe_load(path.read_text(encoding="utf-8"))
        return cls(
            id=d["id"],
            name=d.get("nombre", d["id"]),
            reference_date=str(d["fecha_referencia"]),
            frameworks=d.get("marcos", {}),
            concept_files=[_resolve(p, root) for p in d.get("conceptos", [])],
            rule_patterns=[_resolve(p, root) for p in d.get("reglas", [])],
            severity_file=_resolve(d["gravedad"], root) if d.get("gravedad") else None,
            file=path,
        )


def load_sources(path: str | Path) -> tuple[list[SourceSpec], Path | None]:
    """Lee ``config/sources/*.yaml``: ficheros de entrada de una entidad y sus ajustes."""
    path = Path(path).resolve()
    root = project_root(path.parent)
    d = yaml.safe_load(path.read_text(encoding="utf-8"))
    specs = []
    for s in d.get("fuentes", []):
        opts = dict(s.get("opciones") or {})
        for k in ("layout", "datapoint_map", "fact_map"):
            if k in opts:
                opts[k] = str(_resolve(opts[k], root))
        specs.append(SourceSpec(
            framework=s["marco"],
            loader=s["loader"],
            path=_resolve(s["fichero"], root),
            scale=float(s.get("escala", 1)),
            decimals=s.get("decimales"),
            options=opts,
        ))
    adj = _resolve(d["ajustes"], root) if d.get("ajustes") else None
    return specs, adj


def sha256(path: Path) -> str:
    if path.is_dir():
        return "carpeta"
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


@dataclass
class RunResult:
    exercise: Exercise
    sources: list[SourceSpec]
    store: FactStore
    concepts: ConceptMap
    rules: list[Rule]
    results: list[RuleResult]
    coverage: list[CoverageEntry]
    adjustments: list[Adjustment]
    severity: SeverityConfig
    metadata: dict = field(default_factory=dict)


def run(exercise_path: str | Path, sources_path: str | Path) -> RunResult:
    ex = Exercise.load(exercise_path)
    specs, adj_path = load_sources(sources_path)
    store, coverage = FactStore(), []
    for spec in specs:
        lr = load_source(spec)
        store.extend(lr.facts)
        coverage.extend(lr.coverage)
    cmap = ConceptMap.load(ex.concept_files)
    rules = load_rules(ex.rule_patterns)
    sev = SeverityConfig.load(ex.severity_file)
    adjustments = load_adjustments(adj_path)
    results = evaluate(rules, store, cmap, sev, adjustments)
    meta = {
        "herramienta": f"cuadre {__version__}",
        "ejecutado": dt.datetime.now().isoformat(timespec="seconds"),
        "ejercicio": ex.id,
        "fecha_referencia": ex.reference_date,
        "marcos": ", ".join(f"{k} {v}" for k, v in ex.frameworks.items()),
        "fichero_ejercicio": str(ex.file),
        "fichero_fuentes": str(Path(sources_path).resolve()),
        "tolerancia_eur": sev.tolerance_eur,
        "n_hechos": len(store),
        "n_conflictos": len(store.conflicts),
        "n_reglas": len(rules),
        "hashes": {str(s.path): sha256(s.path) for s in specs},
    }
    return RunResult(ex, specs, store, cmap, rules, results, coverage, adjustments, sev, meta)
