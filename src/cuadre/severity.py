"""Gravedad de un descuadre: escala absoluta y relativa; gana la más grave."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

DEFAULT_LEVELS = ["Leve", "Moderado", "Grave", "Crítico"]


@dataclass
class SeverityConfig:
    tolerance_eur: float = 0.005
    # Límites superiores (exclusivos) de cada grado salvo el último.
    abs_thresholds_eur: list[float] = field(default_factory=lambda: [1_000, 100_000, 1_000_000])
    rel_thresholds: list[float] = field(default_factory=lambda: [0.0001, 0.001, 0.01])
    levels: list[str] = field(default_factory=lambda: list(DEFAULT_LEVELS))

    @classmethod
    def load(cls, path: str | Path | None) -> "SeverityConfig":
        if not path:
            return cls()
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        cfg = cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})
        if len(cfg.abs_thresholds_eur) != len(cfg.levels) - 1 or len(cfg.rel_thresholds) != len(cfg.levels) - 1:
            raise ValueError("Hace falta un umbral menos que grados en cada escala")
        return cfg

    def rank(self, level: str | None) -> int:
        return self.levels.index(level) if level in self.levels else -1

    def grade_abs(self, diff: float) -> str:
        for limit, level in zip(self.abs_thresholds_eur, self.levels):
            if diff < limit:
                return level
        return self.levels[-1]

    def grade_rel(self, diff: float, magnitude: float) -> str | None:
        if magnitude <= 0:
            return None  # sin base no hay gravedad relativa
        ratio = diff / magnitude
        for limit, level in zip(self.rel_thresholds, self.levels):
            if ratio < limit:
                return level
        return self.levels[-1]

    def grade(self, diff: float, magnitude: float, cap: str | None = None) -> tuple[str, str, str | None]:
        """Devuelve (final, absoluto, relativo). ``cap`` limita el grado final."""
        a = self.grade_abs(diff)
        r = self.grade_rel(diff, magnitude)
        final = max([a, r] if r else [a], key=self.rank)
        if cap and self.rank(cap) >= 0 and self.rank(final) > self.rank(cap):
            final = cap
        return final, a, r
