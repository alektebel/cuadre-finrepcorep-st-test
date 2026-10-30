"""Carga del catálogo: conceptos por versión de marco, reglas y ajustes de conciliación.

Las reglas apuntan preferiblemente a *conceptos* (``concept:finrep.total_assets``) y no
a celdas físicas: así un mismo catálogo sirve para COREP 3.3 (12/2024) y 4.x (12/2025),
que tienen códigos de fila/columna distintos. Cada ejercicio declara qué mapas de
conceptos usa.
"""
from __future__ import annotations

import glob
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .expressions import parse_expression
from .model import DatapointRef, Rule, RuleOrigin, RuleType


@dataclass
class Concept:
    name: str
    ref: DatapointRef
    label: str | None = None
    verified: bool = False  # True solo si alguien ha comprobado la celda contra la taxonomía
    origin_file: str | None = None


@dataclass
class ConceptMap:
    concepts: dict[str, Concept] = field(default_factory=dict)

    @classmethod
    def load(cls, paths: list[Path]) -> "ConceptMap":
        cmap = cls()
        for p in paths:
            data = yaml.safe_load(Path(p).read_text(encoding="utf-8")) or {}
            for name, spec in (data.get("conceptos") or {}).items():
                if isinstance(spec, str):
                    spec = {"ref": spec}
                if name in cmap.concepts:
                    raise ValueError(f"Concepto duplicado {name!r} en {p} y {cmap.concepts[name].origin_file}")
                cmap.concepts[name] = Concept(
                    name=name,
                    ref=DatapointRef.parse(spec["ref"]),
                    label=spec.get("etiqueta"),
                    verified=bool(spec.get("verificado", False)),
                    origin_file=str(p),
                )
        return cmap

    def resolve(self, token: str) -> DatapointRef | None:
        """Devuelve la referencia de un token; None si es un concepto sin mapear."""
        if token.startswith("concept:"):
            c = self.concepts.get(token.removeprefix("concept:"))
            return c.ref if c else None
        return DatapointRef.parse(token)


_RULE_KEYS = {
    "id": "id",
    "descripcion": "description",
    "tipo": "type",
    "lhs": "lhs",
    "rhs": "rhs",
    "origen": "origin",
    "fuente": "source",
    "faltantes": "missing",
    "gravedad_maxima": "max_severity",
    "etiquetas": "tags",
    "activa": "enabled",
}


def rule_from_dict(d: dict) -> Rule:
    unknown = set(d) - set(_RULE_KEYS)
    if unknown:
        raise ValueError(f"Regla {d.get('id')!r}: claves desconocidas {sorted(unknown)}")
    kw = {_RULE_KEYS[k]: v for k, v in d.items()}
    kw["type"] = RuleType(kw["type"])
    kw["origin"] = RuleOrigin(kw.get("origin", "propuesta"))
    kw["lhs"] = str(kw.get("lhs", "0"))
    kw["rhs"] = str(kw.get("rhs", "0"))
    kw.setdefault("description", "")
    if kw.get("missing", "error") not in ("error", "zero"):
        raise ValueError(f"Regla {kw['id']}: 'faltantes' debe ser 'error' o 'zero'")
    rule = Rule(**kw)
    parse_expression(rule.lhs)  # falla pronto si la sintaxis es inválida
    parse_expression(rule.rhs)
    if rule.origin == RuleOrigin.NORMATIVA and not (rule.source or {}).get("cita"):
        raise ValueError(f"Regla normativa {rule.id} sin 'fuente.cita': no es trazable")
    return rule


def rule_to_dict(rule: Rule) -> dict:
    inv = {v: k for k, v in _RULE_KEYS.items()}
    out = {}
    for attr, key in inv.items():
        val = getattr(rule, attr)
        if hasattr(val, "value"):
            val = val.value
        if val in (None, [], {}) and key not in ("lhs", "rhs"):
            continue
        out[key] = val
    return out


def load_rules(patterns: list[str | Path]) -> list[Rule]:
    rules: list[Rule] = []
    seen: dict[str, str] = {}
    for pattern in patterns:
        for f in sorted(glob.glob(str(pattern))):
            data = yaml.safe_load(Path(f).read_text(encoding="utf-8")) or {}
            for d in data.get("reglas") or []:
                r = rule_from_dict(d)
                if r.id in seen:
                    raise ValueError(f"Id de regla duplicado {r.id} en {f} y {seen[r.id]}")
                seen[r.id] = f
                r.source = {**(r.source or {}), "_fichero_regla": f}
                rules.append(r)
    return rules


@dataclass
class Adjustment:
    rule_id: str
    amount_eur: float
    explanation: str
    file: str


def load_adjustments(path: str | Path | None) -> list[Adjustment]:
    if not path or not Path(path).exists():
        return []
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    out = []
    for a in data.get("ajustes") or []:
        if not a.get("explicacion"):
            raise ValueError(f"Ajuste de la regla {a.get('regla')} sin explicación: no es trazable")
        out.append(Adjustment(a["regla"], float(a["importe"]), a["explicacion"], str(path)))
    return out
