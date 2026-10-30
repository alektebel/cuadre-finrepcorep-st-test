# Paso 6 — Verificador independiente

Recibes **solo** una regla en el DSL y su cita literal (no el razonamiento de quien la
generó), más las etiquetas de las celdas a las que apunta (inventario).

Pregunta: ¿la cita implica exactamente esta regla?

- `si`: la cita impone esta relación, con estos operandos y este sentido (=, <=, >=).
- `parcial`: la cita apoya la regla pero falta/sobra algún operando, el sentido no es el
  correcto o el perímetro (método, cartera, consolidación) no coincide.
- `no`: la cita no impone esta relación.

Sé estricto: ante la duda, `parcial` con el motivo. Salida: un JSON por regla que valide
contra `schemas/verdict.schema.json`.
