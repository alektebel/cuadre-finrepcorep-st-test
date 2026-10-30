# Paso 4 — Traductor al DSL de reglas

Recibes un candidato (salida del paso 3), el inventario de celdas de las plantillas ST
(`work/inventory_st.json`), los inventarios FINREP/COREP si existen y los mapas de
conceptos (`concepts/*.yaml`).

Objetivo: convertir el candidato en una regla que valide contra
`schemas/rule_draft.schema.json`.

DSL de expresiones: combinación lineal de términos entre llaves, con `+`, `-` y
coeficiente opcional `k*`:

    {concept:st.cet1} + {COREP:C 01.00!r0020c0010} - 2*{concept:x}

Reglas estrictas:
1. Prefiere **conceptos** (`concept:...`) a celdas directas cuando la celda FINREP/COREP
   cambia entre versiones de marco (3.3 vs 4.x). Si el concepto no existe, propónlo en
   `new_concepts` con su `ref` **solo si la celda está en un inventario**.
2. Si un operando no se puede mapear a una celda del inventario, ponlo en `unresolved`.
   No inventes códigos de fila/columna: una regla sin resolver es útil; una inventada
   es dañina.
3. `origen: normativa` exige `fuente.documento`, `fuente.apartado`, `fuente.pagina` y
   `fuente.cita` (la cita literal del candidato).
4. Si el cuadre solo se cumple tras ajustes conocidos (p. ej. reexpresión CRR3 del punto
   de partida), usa `tipo: conciliacion`.
5. `id`: `ST25-<plantilla ST>-<nº>` (estable y legible).
