"""Expresiones lineales de las reglas.

Sintaxis (deliberadamente mínima, solo combinaciones lineales):

    {ST:CSV_CA!r0010c0010} + {concept:finrep.total_assets} - 2*{COREP:C 01.00!r0020c0010}

- Cada término va entre llaves: una referencia directa (``MARCO:PLANTILLA!rFILAcCOL[@Z]``)
  o un concepto (``concept:nombre``) que resuelve el ``ConceptMap`` del ejercicio.
- Coeficiente opcional ``k*`` delante del término.
- ``0`` como expresión vacía (útil en reglas de existencia/no negatividad).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_TERM_RE = re.compile(r"\s*(?P<sign>[+-])?\s*(?:(?P<coef>\d+(?:\.\d+)?)\s*\*\s*)?\{(?P<tok>[^}]+)\}\s*")


@dataclass(frozen=True)
class Term:
    coef: float
    token: str  # "concept:..." o referencia directa

    @property
    def is_concept(self) -> bool:
        return self.token.startswith("concept:")


def parse_expression(expr: str) -> list[Term]:
    expr = expr.strip()
    if expr in ("", "0"):
        return []
    terms: list[Term] = []
    pos = 0
    while pos < len(expr):
        m = _TERM_RE.match(expr, pos)
        if not m or m.end() == pos:
            raise ValueError(f"Expresión no válida cerca de {expr[pos:]!r} en {expr!r}")
        if terms and m["sign"] is None:
            raise ValueError(f"Falta operador antes de {{{m['tok']}}} en {expr!r}")
        sign = -1.0 if m["sign"] == "-" else 1.0
        coef = float(m["coef"]) if m["coef"] else 1.0
        terms.append(Term(sign * coef, m["tok"].strip()))
        pos = m.end()
    return terms


def tokens(expr: str) -> list[str]:
    return [t.token for t in parse_expression(expr)]
