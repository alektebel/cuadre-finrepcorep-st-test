# Paso 3 — Extractor de cuadres

Recibes el texto de varias páginas de un documento EBA del EU-wide stress test
(Template Guidance o Methodological Note), con los párrafos numerados.

Objetivo: localizar **cada afirmación que imponga una relación cuantitativa entre
datos del punto de partida del stress test y datos de FINREP o COREP**, o entre
celdas de las propias plantillas ST.

Son candidatos, por ejemplo, frases del tipo:
- "X shall be consistent with / equal to / reconcile with Y reported in COREP template C ..."
- "the starting point value ... should correspond to FINREP F ... row ... column ..."
- "the sum of ... should equal ..."
- definiciones que remiten a una celda FINREP/COREP concreta ("as defined in C 07.00, column 0200")
  → `kind: definicion_por_referencia`

Reglas estrictas:
1. `quote` es una copia **literal** del texto recibido (se comprueba automáticamente; si no
   aparece tal cual, el candidato se descarta). No parafrasees ni traduzcas.
2. No inventes plantillas, filas ni columnas. Si el texto no las da, deja `hint` vacío.
3. Ignora todo lo relativo a proyecciones (años 1-3 del escenario): solo punto de partida.
4. Si dudas de si una frase impone un cuadre, inclúyela con `confidence: baja`; la
   revisión humana descarta, pero no puede recuperar lo que no extraigas.
5. Salida: JSON Lines, un objeto por línea, que valide contra `schemas/candidate.schema.json`.
   Si no hay candidatos, salida vacía.
