"""Modelo de datos común: referencias a datapoints, hechos, reglas y resultados.

Todo dato, venga de una plantilla ST, de un Excel FINREP/COREP o de una instancia
XBRL, se normaliza a un ``Fact`` identificado por un ``DatapointRef`` y con su
``Source`` (de dónde sale exactamente). La trazabilidad del cuadre depende de que
ningún valor entre al motor sin su origen.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum

_REF_RE = re.compile(
    r"^(?P<fw>[A-Z][A-Z0-9_]*):(?P<tpl>[^!]+)!r(?P<row>[\w.]+)c(?P<col>[\w.]+)(?:@(?P<z>.+))?$"
)


def _norm_code(code: str) -> str:
    """Normaliza códigos de fila/columna: '10' -> '0010', 'r0010' -> '0010'."""
    code = str(code).strip().lstrip("rRcC")
    return code.zfill(4) if code.isdigit() else code


@dataclass(frozen=True, order=True)
class DatapointRef:
    framework: str  # "ST", "FINREP", "COREP", "AJUSTE"
    template: str  # "F 01.01", "C 01.00", "CSV_CA"...
    row: str
    col: str
    z: str | None = None  # eje z / hoja abierta (país, clase de exposición...)

    def __post_init__(self) -> None:
        object.__setattr__(self, "template", self.template.strip())
        object.__setattr__(self, "row", _norm_code(self.row))
        object.__setattr__(self, "col", _norm_code(self.col))

    @classmethod
    def parse(cls, text: str) -> "DatapointRef":
        m = _REF_RE.match(text.strip())
        if not m:
            raise ValueError(
                f"Referencia no válida: {text!r}. Formato: MARCO:PLANTILLA!rFILAcCOL[@Z]"
            )
        return cls(m["fw"], m["tpl"], m["row"], m["col"], m["z"])

    def __str__(self) -> str:
        s = f"{self.framework}:{self.template}!r{self.row}c{self.col}"
        return f"{s}@{self.z}" if self.z else s


@dataclass(frozen=True)
class Source:
    """Origen físico de un hecho."""

    file: str
    sheet: str | None = None  # hoja Excel o tabla CSV
    cell: str | None = None  # "D12" en Excel; contextRef/datapoint en XBRL
    detail: str | None = None

    def __str__(self) -> str:
        parts = [self.file]
        if self.sheet:
            parts.append(self.sheet)
        if self.cell:
            parts.append(self.cell)
        return " › ".join(parts)


@dataclass(frozen=True)
class Fact:
    ref: DatapointRef
    value: float  # ya en EUR
    source: Source
    raw_value: float | None = None  # valor tal cual en el fichero
    scale: float = 1.0  # multiplicador a EUR aplicado (1e6 si venía en millones)
    rounding_eur: float = 0.0  # media unidad del último dígito significativo, en EUR
    label: str | None = None


class RuleType(str, Enum):
    EQ = "igualdad"  # lhs == rhs
    LE = "menor_igual"  # lhs <= rhs
    GE = "mayor_igual"  # lhs >= rhs
    RECON = "conciliacion"  # lhs == rhs + ajustes explicados por la entidad


class RuleOrigin(str, Enum):
    NORMATIVA = "normativa"  # sale literalmente de la guía/metodología EBA
    PLANTILLA = "plantilla"  # control interno embebido en la plantilla ST
    PROPUESTA = "propuesta"  # sugerida (agente o persona), no normativa
    SINTETICA = "sintetica"  # solo para pruebas


@dataclass
class Rule:
    id: str
    description: str
    type: RuleType
    lhs: str  # expresión, ver expressions.py
    rhs: str
    origin: RuleOrigin
    source: dict = field(default_factory=dict)  # documento, apartado, página, cita
    missing: str = "error"  # "error" -> no evaluable; "zero" -> se toma 0
    max_severity: str | None = None  # p. ej. "Leve" para reglas propuestas
    tags: list[str] = field(default_factory=list)
    enabled: bool = True


class Status(str, Enum):
    OK = "OK"
    REDONDEO = "REDONDEO"  # fuera de tolerancia pero explicable por redondeo
    KO = "KO"
    NO_EVALUABLE = "NO_EVALUABLE"


@dataclass
class Operand:
    side: str  # "LHS" / "RHS"
    coef: float
    token: str  # tal cual en la regla (concepto o referencia)
    ref: DatapointRef | None
    fact: Fact | None
    note: str | None = None  # p. ej. "no encontrado", "tomado como 0"

    @property
    def contribution(self) -> float:
        return self.coef * (self.fact.value if self.fact else 0.0)


@dataclass
class RuleResult:
    rule: Rule
    status: Status
    lhs: float | None
    rhs: float | None
    diff_raw: float | None  # lhs - rhs (en desigualdades: cuantía de la violación)
    rounding_noise: float  # suma de redondeos de los operandos, en EUR
    diff_unexplained: float | None  # max(0, |diff| - ruido)
    severity: str  # grado final
    severity_abs: str | None
    severity_rel: str | None
    operands: list[Operand]
    message: str = ""
