# Plan: herramienta de cuadre Stress Test EBA ↔ FINREP/COREP

Estado: **borrador pendiente de validar** (ver §8, preguntas abiertas).
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

## 8. Preguntas abiertas (bloqueantes)

1. ¿Qué ejercicio y qué fecha exacta? ¿ST EU-wide 2025 con punto de partida
   12/2024, u otro con datos a 12/2025?
2. ¿Excel es un requisito duro de *ejecución* (no se puede usar Python) o solo
   de *presentación*?
3. ¿En qué formato llega FINREP/COREP de la entidad: xBRL-XML, xBRL-CSV o
   exports Excel de una herramienta interna (Vermeg/AxiomSL/Wolters Kluwer…)?
4. ¿Ya tienes una lista de reglas de cuadre (por ejemplo, del QA de la
   autoridad competente) o hay que extraerla de la Template Guidance?
5. ¿Qué perímetro: solo punto de partida o también proyecciones?
   (Las proyecciones no tienen contraparte en FINREP/COREP.)
6. ¿Tolerancias: el redondeo a millones de la EBA o un umbral propio?

## 9. Fases

| Fase | Entregable | Criterio de salida |
|---|---|---|
| 0 | Respuestas a §8 + acceso de red a `eba.europa.eu` | Fuentes descargadas y hasheadas en `manifest.yaml` |
| 1 | Inventario de las plantillas ST y extracción de las referencias a FINREP/COREP de la guidance | `rules/SOURCES.md` con cada candidata y su párrafo |
| 2 | Modelo de datos + loaders ST y xBRL-CSV | Tests de carga sobre ficheros reales de la EBA |
| 3 | Motor + catálogo inicial (≈20 reglas de alto valor: total activo, fondos propios, APR total, exposiciones por método) | Test end-to-end sintético en verde |
| 4 | Libro de resultados con drill-down | Un usuario localiza la celda origen de un descuadre en ≤2 clics |
| 5 | Ampliar el catálogo + reglas de conciliación CRR3 + (opcional) libro con fórmulas vivas | Cobertura documentada regla ↔ fuente |
