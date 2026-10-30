from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from ..model import Fact


@dataclass
class SourceSpec:
    """Un fichero de entrada tal y como lo declara ``config/sources/*.yaml``."""

    framework: str  # ST | FINREP | COREP
    loader: str  # excel_annex | excel_declared | xbrl_csv | xbrl_xml
    path: Path
    scale: float = 1.0  # multiplicador a EUR (1e6 si el fichero va en millones)
    decimals: int | None = None  # decimales con los que se reporta; None = exacto
    options: dict = field(default_factory=dict)

    @property
    def rounding_eur(self) -> float:
        return 0.0 if self.decimals is None else 0.5 * self.scale * 10 ** (-self.decimals)


@dataclass
class CoverageEntry:
    """Qué ha leído el loader de cada hoja/tabla: sirve para detectar lecturas erróneas."""

    file: str
    sheet: str
    template: str | None
    z: str | None
    n_rows: int
    n_cols: int
    n_facts: int
    warnings: list[str] = field(default_factory=list)


@dataclass
class LoadResult:
    facts: list[Fact] = field(default_factory=list)
    coverage: list[CoverageEntry] = field(default_factory=list)
