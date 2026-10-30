import pytest

from cuadre.catalog import ConceptMap, rule_from_dict
from cuadre.engine import evaluate_rule
from cuadre.expressions import parse_expression
from cuadre.model import DatapointRef, Fact, Source, Status
from cuadre.severity import SeverityConfig
from cuadre.store import FactStore


def test_ref_roundtrip_and_normalization():
    r = DatapointRef.parse("COREP:C 07.00!r10c0200@CORP")
    assert (r.row, r.col, r.z) == ("0010", "0200", "CORP")
    assert str(r) == "COREP:C 07.00!r0010c0200@CORP"
    with pytest.raises(ValueError):
        DatapointRef.parse("C 07.00 r10")


def test_expression_parsing():
    t = parse_expression("{concept:a} - 2*{ST:T!r1c1} + {concept:b}")
    assert [(x.coef, x.token) for x in t] == [(1, "concept:a"), (-2, "ST:T!r1c1"), (1, "concept:b")]
    assert parse_expression("0") == []
    with pytest.raises(ValueError):
        parse_expression("{a} {b}")


@pytest.mark.parametrize("diff,mag,expected", [
    (500, 1e12, "Leve"),
    (50_000, 1e12, "Moderado"),
    (500_000, 1e12, "Grave"),
    (2_000_000, 1e12, "Crítico"),
    (500, 10_000, "Crítico"),  # 5 % de la partida: gana la escala relativa
])
def test_severity(diff, mag, expected):
    assert SeverityConfig().grade(diff, mag)[0] == expected


def test_severity_cap():
    assert SeverityConfig().grade(5e6, 1e7, cap="Leve")[0] == "Leve"


def _store(*facts):
    s = FactStore()
    for ref, v, rnd in facts:
        s.add(Fact(DatapointRef.parse(ref), v, Source("f.xlsx", "h", "A1"), rounding_eur=rnd))
    return s


def _rule(**kw):
    base = {"id": "T1", "tipo": "igualdad", "lhs": "{ST:A!r1c1}", "rhs": "{FINREP:B!r1c1}", "origen": "sintetica"}
    return rule_from_dict({**base, **kw})


def test_rounding_vs_tolerance():
    sev, cmap = SeverityConfig(), ConceptMap()
    # 3.000 EUR de diferencia con un dato en millones a 2 decimales (redondeo 5.000) -> REDONDEO
    s = _store(("ST:A!r1c1", 100e6, 5000), ("FINREP:B!r1c1", 100e6 + 3000, 0.005))
    assert evaluate_rule(_rule(), s, cmap, sev).status == Status.REDONDEO
    # Sin redondeo declarado, la misma diferencia es KO
    s = _store(("ST:A!r1c1", 100e6, 0), ("FINREP:B!r1c1", 100e6 + 3000, 0))
    r = evaluate_rule(_rule(), s, cmap, sev)
    assert r.status == Status.KO and r.severity == "Moderado"
    # Medio céntimo exacto: OK
    s = _store(("ST:A!r1c1", 1.0, 0), ("FINREP:B!r1c1", 1.005, 0))
    assert evaluate_rule(_rule(), s, cmap, sev).status == Status.OK


def test_inequalities_and_missing():
    sev, cmap = SeverityConfig(), ConceptMap()
    s = _store(("ST:A!r1c1", 10.0, 0), ("FINREP:B!r1c1", 20.0, 0))
    assert evaluate_rule(_rule(tipo="menor_igual"), s, cmap, sev).status == Status.OK
    assert evaluate_rule(_rule(tipo="mayor_igual"), s, cmap, sev).status == Status.KO
    s = _store(("ST:A!r1c1", 0.0, 0))
    assert evaluate_rule(_rule(), s, cmap, sev).status == Status.NO_EVALUABLE
    assert evaluate_rule(_rule(faltantes="zero"), s, cmap, sev).status == Status.OK


def test_normative_rule_requires_quote():
    with pytest.raises(ValueError, match="cita"):
        _rule(origen="normativa", fuente={"documento": "TG"})


def test_conflicting_duplicates_are_reported():
    s = _store(("ST:A!r1c1", 1.0, 0), ("ST:A!r1c1", 2.0, 0))
    assert len(s.conflicts) == 1 and s.get(DatapointRef.parse("ST:A!r1c1")).value == 1.0
