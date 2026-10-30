"""Evaluación de reglas.

Para cada regla se calcula:

- ``diff_raw``: lhs - rhs.
- ``violación``: |diff| en igualdades; la parte que incumple en desigualdades.
- ``rounding_noise``: suma de |coef| x redondeo de cada operando (en EUR). Si un dato
  viene en millones con 2 decimales, su redondeo es 5.000 EUR aunque la tolerancia
  sea 0,005 EUR.
- ``diff_unexplained``: max(0, violación - ruido). La gravedad se calcula sobre esto.

Estados: OK (violación <= tolerancia), REDONDEO (la violación supera la tolerancia
pero cabe en el ruido de redondeo), KO, NO_EVALUABLE (faltan datos).
"""
from __future__ import annotations

from .catalog import Adjustment, ConceptMap
from .expressions import parse_expression
from .model import DatapointRef, Operand, Rule, RuleResult, RuleType, Source, Status, Fact
from .severity import SeverityConfig
from .store import FactStore


def _operands(side: str, expr: str, cmap: ConceptMap, store: FactStore, rule: Rule) -> list[Operand]:
    ops = []
    for term in parse_expression(expr):
        ref = cmap.resolve(term.token)
        if ref is None:
            ops.append(Operand(side, term.coef, term.token, None, None, "concepto sin mapear en este ejercicio"))
            continue
        fact = store.get(ref)
        note = None
        if fact is None:
            note = "tomado como 0 (faltantes: zero)" if rule.missing == "zero" else "dato no encontrado"
        ops.append(Operand(side, term.coef, term.token, ref, fact, note))
    return ops


def evaluate_rule(
    rule: Rule,
    store: FactStore,
    cmap: ConceptMap,
    sev: SeverityConfig,
    adjustments: list[Adjustment] | None = None,
) -> RuleResult:
    ops = _operands("LHS", rule.lhs, cmap, store, rule) + _operands("RHS", rule.rhs, cmap, store, rule)

    if rule.type == RuleType.RECON:
        for a in adjustments or []:
            if a.rule_id == rule.id:
                ref = DatapointRef("AJUSTE", rule.id, "0", "0")
                fact = Fact(ref, a.amount_eur, Source(a.file, detail=a.explanation), label=a.explanation)
                ops.append(Operand("RHS", 1.0, f"ajuste:{a.explanation}", ref, fact))

    unresolved = [o for o in ops if o.ref is None]
    missing = [o for o in ops if o.ref is not None and o.fact is None]
    if unresolved or (missing and rule.missing == "error"):
        names = ", ".join(o.token for o in unresolved + missing)
        return RuleResult(rule, Status.NO_EVALUABLE, None, None, None, 0.0, None, "—", None, None, ops,
                          f"Faltan operandos: {names}")

    lhs = sum(o.contribution for o in ops if o.side == "LHS")
    rhs = sum(o.contribution for o in ops if o.side == "RHS")
    diff = lhs - rhs
    if rule.type == RuleType.LE:
        violation = max(0.0, diff)
    elif rule.type == RuleType.GE:
        violation = max(0.0, -diff)
    else:
        violation = abs(diff)

    noise = sum(abs(o.coef) * o.fact.rounding_eur for o in ops if o.fact)
    unexplained = max(0.0, violation - noise)

    if violation <= sev.tolerance_eur:
        return RuleResult(rule, Status.OK, lhs, rhs, diff, noise, 0.0, "—", None, None, ops)
    if unexplained <= sev.tolerance_eur:
        return RuleResult(rule, Status.REDONDEO, lhs, rhs, diff, noise, unexplained, "—", None, None, ops,
                          f"Diferencia de {violation:,.2f} EUR dentro del redondeo declarado ({noise:,.2f} EUR)")
    final, a, r = sev.grade(unexplained, max(abs(lhs), abs(rhs)), cap=rule.max_severity)
    return RuleResult(rule, Status.KO, lhs, rhs, diff, noise, unexplained, final, a, r, ops)


def evaluate(rules, store, cmap, sev, adjustments=None) -> list[RuleResult]:
    return [evaluate_rule(r, store, cmap, sev, adjustments) for r in rules if r.enabled]
