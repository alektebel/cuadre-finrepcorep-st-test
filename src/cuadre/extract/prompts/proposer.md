# Paso P — Proponedor de cuadres no normativos

Recibes el inventario de celdas ST y los inventarios FINREP/COREP, junto con las
propuestas deterministas por similitud de etiquetas (`work/proposals_labels.jsonl`).

Objetivo: proponer cuadres **con sentido económico** que la guía no exige
explícitamente: mismo concepto en dos plantillas, totales = suma de desgloses,
coherencias de signo o de orden (p. ej. exposición neta <= exposición bruta).

Reglas:
1. `origen: propuesta`, `gravedad_maxima: Leve` siempre.
2. En `descripcion`, la justificación económica en una frase. Si el cuadre puede no
   cumplirse legítimamente (diferencias de perímetro, de definición, de método), dilo.
3. Solo celdas presentes en los inventarios.
4. Salida: JSON Lines contra `schemas/rule_draft.schema.json` con
   `candidate_id: "PROP-<n>"`.
