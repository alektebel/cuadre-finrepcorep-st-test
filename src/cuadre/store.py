"""Almacén de hechos con detección de duplicados en conflicto."""
from __future__ import annotations

from dataclasses import dataclass, field

from .model import DatapointRef, Fact


@dataclass
class FactStore:
    facts: dict[DatapointRef, Fact] = field(default_factory=dict)
    conflicts: list[tuple[Fact, Fact]] = field(default_factory=list)

    def add(self, fact: Fact) -> None:
        prev = self.facts.get(fact.ref)
        if prev is None:
            self.facts[fact.ref] = fact
        elif abs(prev.value - fact.value) > max(prev.rounding_eur, fact.rounding_eur, 1e-9):
            # Se conserva el primero y se informa: decidir cuál es correcto no es tarea del motor.
            self.conflicts.append((prev, fact))

    def extend(self, facts) -> None:
        for f in facts:
            self.add(f)

    def get(self, ref: DatapointRef) -> Fact | None:
        return self.facts.get(ref)

    def by_framework(self, framework: str) -> list[Fact]:
        return [f for f in self.facts.values() if f.ref.framework == framework]

    def __len__(self) -> int:
        return len(self.facts)
