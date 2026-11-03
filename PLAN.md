# Plan: herramienta de cuadre Stress Test EBA ↔ FINREP/COREP

Estado: **v3, esqueleto implementado** (motor, loaders, informe Excel, app Streamlit y
flujo de extracción con contratos para agentes externos), validado con datos sintéticos.
Pendiente: documentos EBA reales (fase 0) para poblar conceptos y reglas.
Ninguna fuente de la EBA se ha descargado todavía: el dominio `eba.europa.eu`
está bloqueado por la política de red del entorno donde se redactó este plan.

---

## 1. Objetivo

Dada una entidad, comprobar que los datos de punto de partida que declara en las
plantillas del stress test EU-wide de la EBA cuadran con lo que ha remitido en
FINREP y COREP a la misma fecha. Para cada descuadre se debe poder ver:

- qué regla falla (identificador y fuente normativa: documento, apartado, página);
- qué celdas intervienen en cada lado (plantilla ST + fila/columna; plantilla
  FINREP/COREP + fila/columna/eje z);
- el valor de cada operando, el resultado esperado, la diferencia y la tolerancia
  aplicada;
- de qué fichero de entrada sale cada valor (ruta, hoja/CSV, celda o datapoint).

## 2. Riesgos de base que condicionan el diseño

Hay que resolverlos antes de escribir código; cualquiera de ellos puede invalidar
el enfoque.

1. **Fecha de referencia.** El ejercicio EU-wide 2025 usa como punto de partida
   el **31/12/2024** (horizonte 2025–2027; resultados publicados el 1/8/2025).
   "Stress test 2025-12" no corresponde a ningún ejercicio EU-wide conocido
   (el siguiente es 2027). Hay que confirmar si se refiere al ejercicio 2025
   (dato a 12/2024) o a otra cosa (un ejercicio interno o de la autoridad
   nacional con datos a 12/2025).
2. **Ruptura CRR3.** En 2025 la EBA exigió reexpresar exposiciones y APR
   (REA) de riesgo de crédito del punto de partida con las reglas y clases de
   exposición de CRR3. COREP a 31/12/2024 se remitió con el marco 3.3
   (pre-CRR3; el marco 4.0 aplica desde 03/2025). Consecuencia: **una parte de
   las igualdades ST = COREP no se cumple por construcción**. La herramienta
   necesita reglas de tipo "conciliación con partida explicativa", no solo
   igualdades estrictas.
3. **Libro bancario / de negociación.** El reparto que pide la metodología no
   existe tal cual en FINREP, así que algunas reglas solo pueden ser
   desigualdades o cotas, no igualdades.
4. **Origen de las reglas.** Hay que verificar que exista una lista publicada
   y exhaustiva de reglas ST ↔ FINREP/COREP. Lo seguro es que la *Template
   Guidance* y la *Methodological Note* remiten a celdas de FINREP/COREP en
   varias definiciones, y que las plantillas Excel traen controles internos.
   Que exista un fichero formal de validaciones cruzadas con FINREP/COREP es
   **incierto (confianza media-baja)**. Si no existe, el catálogo de reglas se
   tendrá que construir a mano a partir de esos textos, y ese es el grueso del
   trabajo.
5. **Datos de ejemplo.** Las instancias de ejemplo de la taxonomía EBA son
   sintéticas e independientes entre sí, y la EBA solo publica las plantillas
   ST vacías. Juntarlas daría cientos de descuadres sin sentido. Para probar la
   herramienta hay que **generar un juego sintético coherente** con descuadres
   introducidos a propósito y conocidos de antemano (ver §6).
6. **Formato FINREP/COREP.** Las entidades remiten en XBRL (xBRL-XML; xBRL-CSV
   obligatorio desde la fecha de referencia 31/03/2026). Excel no lee XBRL de
   forma nativa, así que hace falta una capa de extracción.

## 3. Decisión de arquitectura: Excel puro vs. híbrido

| Criterio | Excel puro (fórmulas/Power Query/VBA) | Híbrido: motor Python + salida Excel |
|---|---|---|
| Leer XBRL / xBRL-CSV | Power Query puede leer xBRL-CSV (CSV + JSON); XBRL-XML es frágil | Bibliotecas maduras (`arelle`, `pandas`) |
| Catálogo de cientos de reglas | Fórmulas por regla: difíciles de versionar y auditar | Reglas como datos (YAML/CSV), probadas automáticamente |
| Trazabilidad | Posible con hipervínculos, pero se rompe al cambiar el layout | Se genera en cada ejecución (hipervínculos a celdas incluidos) |
| Control de versiones / revisión | Binario, difícil de comparar con `diff` | Texto plano en git |
| Uso por negocio sin instalar nada | Sí | Hay que ejecutar un script, o distribuir un `.exe` o una macro que lo lance |

**Recomendación:** modelo híbrido.
- **Fuente de verdad:** catálogo de reglas en texto (`rules/*.yaml`), versionado.
- **Motor:** Python, que carga las plantillas ST (xlsx/csv) y FINREP/COREP
  (xBRL-CSV, xBRL-XML o export Excel), normaliza todo a un modelo común de
  datapoints y evalúa las reglas.
- **Salida:** un libro Excel de resultados con hipervínculos y drill-down,
  pensado para el usuario de negocio.
- **Alternativa "Excel-first"** si es requisito duro: generar desde el mismo
  catálogo un libro con fórmulas reales (`=ST!E12-F0101!C010`) para que el
  usuario vea la fórmula viva. Se puede añadir más tarde sin cambiar el núcleo.

Argumento en contra de esta recomendación: si el usuario final solo puede usar
Excel corporativo sin Python, el híbrido no se puede desplegar y el coste pasa
a mantener VBA/Power Query. Por eso es la pregunta abierta 2.

## 4. Modelo de datos

Clave común para cualquier dato:

```
DatapointRef = (marco, plantilla, fila, columna, z?)   # p. ej. ("FINREP", "F 01.01", "0010", "0010", None)
                                                       #      ("ST2025", "CSV_CR_SCEN", "r0030", "c0040", "IRB|Corporates")
Fact = (DatapointRef, valor, unidad, escala, origen{fichero, hoja/csv, celda|contextRef})
```

Regla (YAML):

```yaml
id: ST-FIN-001
descripcion: Total activo ST = Total activo FINREP
fuente: {documento: "2025 EU-wide ST Template Guidance", apartado: "x.y", pagina: n}
tipo: igualdad            # igualdad | desigualdad | conciliacion | existencia
lhs: "ST2025:CSV_xxx!r0010c0010"
rhs: "FINREP:F 01.01!r0380c0010"
escala: {lhs: 1e6, rhs: 1}  # ST en millones, FINREP en unidades
tolerancia: {absoluta: 0.5e6, relativa: null}
partidas_conciliacion: []  # para reglas CRR3: ajustes explicados por la entidad
severidad: bloqueante      # bloqueante | aviso
```

Cada regla evaluada guarda los operandos resueltos con su origen, y de ahí sale
la trazabilidad.

## 5. Estructura del repositorio

```
cuadre-finrepcorep-st-test/
├── README.md
├── PLAN.md
├── pyproject.toml
├── data/
│   ├── raw/                      # descargas originales, sin tocar (gitignored + manifest)
│   │   ├── eba_st2025/           # plantillas, template guidance, methodological note
│   │   └── eba_taxonomy/         # taxonomía + sample instances FINREP/COREP
│   ├── manifest.yaml             # URL, fecha de descarga y SHA-256 de cada fichero
│   └── synthetic/                # juego de prueba coherente generado (§6)
├── rules/
│   ├── st2025_finrep.yaml
│   ├── st2025_corep.yaml
│   ├── st2025_internal.yaml      # controles internos de las propias plantillas ST
│   └── SOURCES.md                # de qué párrafo sale cada regla
├── src/cuadre/
│   ├── download.py               # descarga + verificación de hash
│   ├── loaders/
│   │   ├── st_templates.py       # xlsx/csv EBA ST -> Facts
│   │   ├── xbrl_csv.py           # FINREP/COREP xBRL-CSV -> Facts
│   │   ├── xbrl_xml.py           # vía arelle (opcional)
│   │   └── excel_export.py       # por si la entidad tiene exports Excel propios
│   ├── model.py                  # DatapointRef, Fact, Rule, Result
│   ├── rules.py                  # parser + validación de esquema del catálogo
│   ├── engine.py                 # evaluación, tolerancias, escalas
│   └── report/
│       ├── excel.py              # libro de resultados con drill-down
│       └── formulas.py           # (fase 5) libro con fórmulas vivas
├── tests/
│   ├── test_loaders.py
│   ├── test_engine.py
│   └── test_end_to_end.py        # sintético con descuadres conocidos
└── docs/
    ├── metodologia.md            # riesgos §2 y cómo los trata cada regla
    └── uso.md
```

## 6. Datos de prueba

1. Descargar las plantillas ST 2025 (versión final) y la documentación.
2. Descargar las instancias de ejemplo FINREP/COREP del paquete de taxonomía
   del marco que aplique a la fecha de referencia (3.3 para 12/2024; 4.x si
   la fecha es 12/2025).
3. Generar `data/synthetic/` con este proceso:
   - partir de una instancia FINREP/COREP de ejemplo;
   - rellenar las plantillas ST *derivándolas* de esa instancia aplicando las
     reglas al revés, de modo que todo cuadre por construcción;
   - introducir N descuadres con semilla fija y guardarlos en
     `expected_breaks.yaml`.
4. Test de aceptación: la herramienta detecta exactamente los N descuadres
   esperados, sin falsos positivos ni negativos, y cada uno apunta a la celda
   correcta.

Sin el paso 3, los ejemplos no sirven para comprobar que la herramienta acierta:
solo demostrarían que no se rompe.

## 7. Salida (libro de resultados)

| Hoja | Contenido |
|---|---|
| `Resumen` | Nº de reglas evaluadas, OK/KO/no evaluables, por plantilla y severidad |
| `Resultados` | Una fila por regla: id, descripción, LHS, RHS, diferencia, tolerancia, estado, enlace al detalle |
| `Detalle` | Una fila por operando: regla, lado, DatapointRef, valor, escala, fichero, hoja, celda/contexto, hipervínculo |
| `No_evaluables` | Reglas con datapoints ausentes, y el motivo |
| `ST_*`, `FINREP_*`, `COREP_*` | Copia de los datos de entrada con las celdas implicadas en descuadres resaltadas |
| `Metadatos` | Versión del catálogo, hashes de entrada, fecha de ejecución, versión de la herramienta |

## 8. Decisiones tomadas (respuestas del usuario, 2026-10-30)

| Tema | Decisión |
|---|---|
| Ejercicio | EU-wide stress test. La herramienta se parametriza por **ejercicio + fecha de referencia** y soporta dos configuraciones: `st2025@2024-12-31` (oficial) y `st2025@2025-12-31` (plantillas 2025 con datos de 12/2025). |
| Plataforma | No tiene que ser Excel. Motor Python, revisión en una app local y exportación a Excel trazable. |
| Formato FINREP/COREP | Prioridad: Excel. Además, xBRL-CSV y xBRL-XML. |
| Origen de los cuadres | (a) Extraídos de la guía EBA mediante un flujo de agentes; (b) cuadres propuestos a partir de las celdas, separados de los normativos. |
| Perímetro | Solo el punto de partida. Las proyecciones quedan fuera. |
| Tolerancia | 0,005 EUR, con grados de gravedad según el importe del descuadre. |

### 8.1 Riesgos de estas decisiones

- **Tolerancia de 0,005 EUR frente a datos en millones.** Creo (confianza media;
  pendiente de verificar en la Template Guidance) que las plantillas ST piden los
  importes en millones de EUR. Si una cifra viene redondeada a millones, solo el
  redondeo puede dar hasta ±500.000 EUR de diferencia, y todas las reglas saldrían KO.
  Tratamiento:
  1. Pasar todo a EUR antes de comparar.
  2. Guardar la precisión declarada de cada operando (decimales en XBRL; inferida en Excel).
  3. Mostrar dos cifras:
     - `dif_bruta`;
     - `dif_no_explicada = max(0, |dif| - ruido_redondeo)`, donde `ruido_redondeo`
       es la suma de medias unidades de la precisión de cada operando.
  4. La tolerancia de 0,005 EUR y los grados de gravedad se aplican a
     `dif_no_explicada`. La regla se marca KO si `dif_no_explicada > 0,005 EUR`.
     Si `dif_bruta` supera 0,005 EUR pero el exceso cabe en el redondeo, se
     marca como "redondeo". Así no se oculta el descuadre.
- **Grados de gravedad.** Los umbrales son configurables (`config/severity.yaml`).
  Por defecto se combina una escala absoluta y otra relativa, y gana la más grave:

  | Grado | Absoluto (EUR) | Relativo (dif / max(\|LHS\|,\|RHS\|)) |
  |---|---|---|
  | OK | ≤ 0,005 | — |
  | Redondeo | dentro del ruido de redondeo | — |
  | Leve | < 1.000 | < 0,01 % |
  | Moderado | < 100.000 | < 0,1 % |
  | Grave | < 1.000.000 | < 1 % |
  | Crítico | ≥ 1.000.000 | ≥ 1 % |

  Motivo de la parte relativa: 900.000 EUR de diferencia sobre un total de
  activo de 500.000 M es ruido, pero sobre una partida de 2 M es un error de
  fondo. Con una escala solo absoluta se ordenan mal.
- **FINREP/COREP en Excel.** No hay un layout estándar: cada herramienta
  (AxiomSL, Vermeg, Wolters Kluwer, exportación propia) saca un Excel distinto.
  Se implementan dos loaders:
  1. **Layout de anexo EBA:** detecta los códigos de fila y columna (`0010`,
     `0020`…) en las cabeceras de cada hoja. Cubre la mayoría de exports.
  2. **Layout declarado:** un YAML por fuente indica hoja, rango y qué celda
     es cada datapoint.

  El riesgo es que el loader 1 se equivoque al detectar las cabeceras sin
  avisar. Por eso cada carga genera un informe de cobertura: datapoints
  esperados, encontrados y ambiguos.
- **Dos marcos de reporting.** Con CRR3 cambian los códigos de fila y columna de
  COREP: 12/2024 se reportó con el marco 3.3 y 12/2025 con el 4.x. Las reglas
  no apuntan a celdas físicas, sino a **conceptos** (`corep.own_funds.cet1`,
  `finrep.total_assets`…). Una capa de mapeo por versión traduce cada concepto
  a su celda. Sin esa capa, el catálogo tendría que duplicarse por ejercicio.

### 8.2 Interfaz: app local + Excel

- **App local:** Python con Streamlit (o FastAPI + HTML estático). Muestra una
  tabla de checks filtrable por gravedad, plantilla y estado. Al pulsar un check
  se ve el detalle de operandos, con los valores, la celda de origen, el
  fichero y la cita normativa. Además hay un visor de cada plantilla de entrada
  con las celdas implicadas resaltadas.
- **Angular, descartado para la v1:** obliga a mantener dos lenguajes y una
  cadena de build de Node, y añade una API REST para un único usuario en local.
  No aporta nada que Streamlit no dé para este caso. Solo tendría sentido si la
  herramienta pasa a ser multiusuario o se despliega en un servidor.
- **Export Excel (§7):** el mismo resultado, con hipervínculos y un libro
  opcional de fórmulas vivas. Es el entregable que se archiva y se envía
  como evidencia.

### 8.3 Flujo de agentes para extraer los cuadres

Requisito previo: tener los documentos. Hay dos vías: permitir `eba.europa.eu`
en la red del entorno, o subirlos a `data/raw/eba_st2025/`. Sin ellos el flujo
**no se ejecuta**: un modelo que "recuerda" la guía produce reglas sin fuente
verificable.

```
 PDF guía + methodological note           Plantillas ST (xlsx)
        │                                        │
  [1] ingest (determinista)               [2] inventario (determinista)
  texto por página + nº párrafo           celda → (plantilla, fila, col, etiqueta,
        │                                  fórmulas/controles internos)
        ▼                                        │
  [3] Extractor (Sonnet 5.5, por bloques) ◄──────┘
  → candidatos JSON: {cita literal, página, operandos en texto, tipo}
        ▼
  [4] Traductor (Sonnet 5.5)
  → regla en DSL (rules/*.yaml) con referencias a conceptos/celdas del inventario
        ▼
  [5] Verificación determinista
  - la cita aparece literal en la página indicada
  - todas las celdas/conceptos existen en el inventario
  - el YAML valida contra el esquema
        ▼
  [6] Verificador (otro modelo, sin ver el razonamiento del extractor)
  "¿la cita implica exactamente esta regla?" → sí / no / parcial + motivo
        ▼
  [7] Revisión humana (hoja Excel / pantalla de la app): aceptar, editar, rechazar
```

- Los controles internos que ya vienen como fórmulas en las plantillas ST se
  extraen en el paso [2] con código. No necesitan LLM, que en este caso solo
  aportaría error.
- **Cuadres propuestos (no normativos):**
  - Un agente [P] cruza etiquetas y definiciones de las celdas ST con las
    etiquetas de las celdas FINREP/COREP. También busca identidades
    aritméticas evidentes: totales = suma de desgloses, mismo concepto en dos
    plantillas ST.
  - Estas reglas llevan `origen: propuesta` y gravedad máxima "aviso" hasta
    que alguien las acepte.
  - En la app se ven en una pestaña aparte. Si se mezclan con las normativas,
    se pierde la trazabilidad que se busca.
- **Ejecución:**
  - La primera pasada se hace dentro de la sesión, con subagentes Sonnet 5.5.
  - Después se convierte en un script reproducible (`tools/extract_rules/`)
    con la API de Anthropic. La guía cambia en cada ejercicio (bienal) y el
    flujo tiene que poder repetirse y auditarse: prompts versionados y
    salidas guardadas.

## 9. Fases

| Fase | Entregable | Criterio de salida |
|---|---|---|
| 0 | Acceso a los documentos EBA (red o subida manual) | Fuentes descargadas y hasheadas en `manifest.yaml` |
| 1 | Flujo de agentes §8.3 (pasos 1–7) + cuadres propuestos | Catálogo revisado; cada regla normativa con cita verificada literalmente |
| 2 | Modelo de datos + capa de conceptos por versión de marco + loaders (ST, Excel anexo EBA, Excel declarado, xBRL-CSV, xBRL-XML) | Tests de carga sobre ficheros reales de la EBA |
| 3 | Motor + catálogo inicial (≈20 reglas de alto valor: total activo, fondos propios, APR total, exposiciones por método) | Test end-to-end sintético en verde |
| 4 | App local (Streamlit) + libro de resultados con drill-down | Un usuario localiza la celda origen de un descuadre en ≤2 clics |
| 5 | Ampliar el catálogo + reglas de conciliación CRR3 + (opcional) libro con fórmulas vivas | Cobertura documentada regla ↔ fuente |

### Estado de las fases (2026-10-30)

| Fase | Estado |
|---|---|
| 0 | Bloqueada: falta acceso a `eba.europa.eu` o subir los documentos |
| 1 | Esqueleto hecho (`cuadre-extract`: ingest, inventory, tasks, review, propose, export). Los agentes los pone el usuario |
| 2 | Hecho: modelo, conceptos por versión, loaders `excel_annex`, `excel_declared`, `xbrl_csv`, `xbrl_xml`. **xBRL sin validar contra paquetes reales** |
| 3 | Motor hecho; el catálogo real está vacío. Test end-to-end sintético en verde |
| 4 | Hecho: app Streamlit e informe Excel con drill-down |
| 5 | Pendiente |

### Cambio de objetivo (2026-11-03): entregable final = libro Excel

Decisiones del usuario:
- VBA permitido (`.xlsm`).
- Excel autónomo generado desde el catálogo YAML, sin Python en el puesto.
- Ficheros elegidos con un selector local.

Implementación:
- `cuadre excel` genera el libro con fórmulas vivas.
- `Cuadre.bas` carga los ficheros.
- `tools/build_xlsm.ps1` empaqueta el `.xlsm`.

El motor Python sigue siendo la referencia: un test exige que el libro recalculado dé
lo mismo que el motor. La app Streamlit pasa a ser una herramienta de mantenimiento
(revisar reglas extraídas y depurar), no el producto final.

### Fuentes indicadas por el usuario (2026-11-03)

Registradas en `data/manifest.yaml`; se descargan con `python tools/download_sources.py`:

- **EBA/GL/2018/04**: *Guidelines on institutions' stress testing*. Tratan del programa
  interno de stress testing de cada entidad.
- **EBA/GL/2021/10**: según fuentes secundarias, *Guidelines on stress tests of deposit
  guarantee schemes* (revisadas).

Ninguna de las dos parece contener cuadres de plantillas ST ↔ FINREP/COREP a nivel de
celda. Hay que confirmarlo leyendo los PDF; si se confirma, el extractor sacaría pocos o
ningún candidato de ellas. No se han podido descargar: la red del entorno bloquea
`eba.europa.eu` y las réplicas probadas.
