# cuadre-finrepcorep-st-test

Cuadre trazable entre las plantillas del **EU-wide stress test (EBA)** y los reportes
**FINREP/COREP** de una entidad, solo para el punto de partida. Cada descuadre se puede
seguir hasta la regla, su cita normativa, cada operando y la celda del fichero de origen.

Diseño y decisiones: [`PLAN.md`](PLAN.md).

## El libro Excel (entregable principal)

Un libro autónomo: quien lo usa no necesita Python. Se cargan con botones los Excel de
ST, FINREP y COREP y las hojas de cuadre se recalculan solas.

```bash
# 1. Generar el libro a partir del catálogo (reglas + conceptos + gravedad)
cuadre excel --ejercicio config/exercises/st2025_2024-12.yaml --salida out/cuadre.xlsx
# 2. En Windows con Excel: añadir macros y botones -> out/cuadre.xlsm
powershell -ExecutionPolicy Bypass -File tools/build_xlsm.ps1 -Xlsx out/cuadre.xlsx
```

Sin PowerShell: abre el `.xlsx`, `Alt+F11` → Archivo → Importar →
`src/cuadre/excel_book/vba/Cuadre.bas`, guarda como `.xlsm` y ejecuta la macro
`CrearBotones` una vez.

### Requisitos: Excel de escritorio en Windows (Microsoft 365 / Office 2016+)

- **No funciona en Excel para la web** (navegador, OneDrive/SharePoint/Teams), porque
  ahí no se ejecuta VBA. Ábrelo con «Abrir en la aplicación de escritorio».
- **Macros bloqueadas al descargar.** Microsoft 365 bloquea las macros de cualquier
  `.xlsm` que venga de internet, del correo o de GitHub (barra roja «Microsoft ha
  bloqueado la ejecución de macros»). Antes de abrirlo: clic derecho en el fichero →
  *Propiedades* → marcar **Desbloquear** → Aceptar. Otra opción es guardarlo en una
  *ubicación de confianza* (Archivo → Opciones → Centro de confianza).
- **Para `build_xlsm.ps1`**, una sola vez: Archivo → Opciones → Centro de confianza →
  Configuración → Configuración de macros → «Confiar en el acceso al modelo de objetos de
  proyectos de VBA».
- Las fórmulas se guardan en sintaxis inglesa y Excel en español las muestra traducidas
  (`SUMAR.SI.CONJUNTO`, `;`…). No hace falta hacer nada.

Demo con datos sintéticos ya cargados: [`dist/cuadre_sintetico.xlsx`](dist/cuadre_sintetico.xlsx)
(los enlaces «abrir» funcionan si se mantiene la estructura del repositorio).

| Hoja | Contenido |
|---|---|
| Inicio | Instrucciones y botones Cargar ST / FINREP / COREP, Recalcular, Borrar |
| Resumen | Recuento por estado, gravedad y origen; datos cargados y duplicados |
| Checks | Una regla por fila: LHS, RHS, diferencia, redondeo admitido, no explicada, estado, gravedad, cita, → operandos |
| Operandos | Un término por fila: clave, valor, redondeo, fichero/hoja/celda, → dato, abrir origen, ← check |
| Datos | Un hecho por fila (lo escribe la macro); marca en rojo los implicados en KO y señala duplicados |
| Carga | Registro por hoja leída: filas/columnas detectadas, nº de datos, avisos |
| Mapeo_hojas | Hoja → plantilla / eje Z / omitir, para exports cuyo nombre de hoja no es la plantilla |
| Conceptos, Ajustes, Parametros | Editables: mapa de conceptos, partidas de conciliación, tolerancia y umbrales |

Las fórmulas solo usan funciones de Excel 2010+ (SUMIFS, COUNTIFS, INDEX/MATCH). Un test
recalcula el libro con LibreOffice y comprueba que da **exactamente** los mismos estados,
gravedades e importes que el motor Python.

**Límites del libro:**
- La macro VBA no se ha ejecutado en Excel; solo se verifica que es coherente con el
  generador (hojas, columnas, macros). Hay que probarla en Windows antes de usarla en real.
- La macro solo lee Excel con layout de anexo EBA. Para xBRL, o para exports con layout
  propio, usa `cuadre excel --precargar config/sources/<entidad>.yaml`, que rellena
  `Datos` con los loaders Python.
- Si una clave está duplicada en `Datos`, cuenta la primera (igual que el motor) y se
  marca como `DUPLICADO`.

## Instalación

```bash
pip install -e ".[app,dev]"
```

## Ejercicio 2027 (borrador EBA de 11/06/2026)

Las plantillas oficiales ST traen una columna **«COREP CODE / COMPUTATION»** con la celda
COREP de cada fila. De ella sale el catálogo de cuadres de punto de partida sin LLM:

```bash
cuadre-extract st-plantillas --xlsx "data/raw/eba_st2027/2027 EU-wide stress test - Draft Templates.xlsx" \
                             --rules-dir rules/st2027 [--controles-internos]
cuadre excel --ejercicio config/exercises/st2027_2026-12.yaml --salida out/cuadre_st2027.xlsx
```

Resultado actual: **2.489 reglas** ST = COREP (`CSV_CAP`, `CSV_REA_SUM` y las tres hojas de
REA de riesgo de crédito). El detalle está en [`rules/st2027/INFORME_EXTRACCION.md`](rules/st2027/INFORME_EXTRACCION.md)
y el libro, en [`dist/cuadre_st2027.xlsx`](dist/cuadre_st2027.xlsx).

- **Formato ST:** la macro y el loader `st_eba` leen el layout oficial (`RowNum` + columnas
  numeradas). Referencia: `ST:<hoja>!r<RowNum>c<nº columna>`.
- **Eje Z geográfico:** en riesgo de crédito las celdas COREP van con Z
  `GEO=<Geographical breakdown>`, por ejemplo `COREP:C 09.01!r0010c0075@GEO=Total`. Hay que
  mapear cada hoja por país del export COREP a ese Z (hoja `Mapeo_hojas` o `opciones.sheets`).
- **Alternativas:** las filas «C 09.02 … or C 09.05 …» generan dos reglas con la etiqueta
  `alternativa`. La que la entidad no reporte quedará `NO_EVALUABLE`.

## Motor Python y app local (juego sintético)

```bash
python -m cuadre.synthetic data/synthetic          # regenera datos de prueba (ya versionados)
cuadre run --ejercicio config/exercises/synthetic.yaml \
           --fuentes config/sources/synthetic.yaml --salida out/informe.xlsx
streamlit run app/streamlit_app.py                 # app local
pytest
```

Las plantillas `SYN_*` son **inventadas**: sirven para probar la herramienta, no para
representar plantillas EBA.

## Piezas

| Pieza | Dónde | Qué hace |
|---|---|---|
| Modelo | `src/cuadre/model.py` | `DatapointRef` (`MARCO:PLANTILLA!rFILAcCOL[@Z]`), `Fact` con origen, `Rule`, `RuleResult` |
| Conceptos | `concepts/*.yaml` | Nombre estable → celda, **por versión de marco** (COREP 3.3 vs 4.x) |
| Reglas | `rules/**/*.yaml` | DSL lineal: `{concept:st.cet1} - {COREP:C 01.00!r0020c0010}` · tipos `igualdad`, `menor_igual`, `mayor_igual`, `conciliacion` |
| Loaders | `src/cuadre/loaders/` | `excel_annex` (detecta códigos de fila/columna), `excel_declared` (layout YAML), `xbrl_csv`, `xbrl_xml` |
| Motor | `src/cuadre/engine.py` | Tolerancia 0,005 EUR **sobre la diferencia no explicada por redondeo** |
| Gravedad | `config/severity.yaml` | Escala absoluta y relativa; gana la peor. Las reglas `propuesta` se limitan con `gravedad_maxima` |
| Ajustes | `ajustes:` en las fuentes | Partidas de conciliación explicadas (p. ej. reexpresión CRR3) para reglas `conciliacion` |
| Informe | `src/cuadre/report/excel.py` | Resumen → Resultados → Detalle → celda de origen, todo con hipervínculos |
| App | `app/streamlit_app.py` | Checks filtrables, detalle por operando, visor de plantillas con celdas marcadas, cobertura de carga, revisión de reglas |
| Extracción | `src/cuadre/extract/` | Pasos deterministas y contratos JSON para tus agentes (ver abajo) |

## Estados de un check

- `OK`: diferencia ≤ 0,005 EUR.
- `REDONDEO`: la diferencia supera la tolerancia pero cabe en el redondeo declarado de los
  operandos (un dato en millones con 2 decimales arrastra ±5.000 EUR).
- `KO`: con gravedad Leve / Moderado / Grave / Crítico.
- `NO_EVALUABLE`: falta un dato o un concepto no está mapeado en ese ejercicio.

## Flujo de extracción de reglas con agentes

```
cuadre-extract ingest --pdf <guía.pdf> --doc TG2025                 # [1] texto por página/párrafo
cuadre-extract inventory --xlsx <plantillas ST> --marco ST --out work/inventory_st.json   # [2]
cuadre-extract tasks extract     # work/tasks/extract.jsonl  → tus agentes → work/agent_out/extract/*.jsonl
cuadre-extract tasks translate   # work/tasks/translate.jsonl → tus agentes → work/agent_out/translate/*.jsonl
cuadre-extract propose           # propuestas deterministas por similitud de etiquetas
cuadre-extract review            # checks deterministas → work/review.json
cuadre-extract tasks verify      # work/tasks/verify.jsonl   → tus agentes → work/agent_out/verify/*.jsonl
cuadre-extract review            # incorpora veredictos (conserva decisiones previas)
# revisión humana en la app, pestaña «Revisión de reglas»
cuadre-extract export --rules-dir rules/st2025
```

Cada tarea JSONL lleva `prompt` (en `src/cuadre/extract/prompts/`), `output_schema` (en
`src/cuadre/extract/schemas/`), `input` y la ruta `output` donde el agente debe escribir.
Checks deterministas que no dependen del agente:

- la cita aparece **literalmente** en la página indicada;
- todas las celdas existen en el inventario de plantillas;
- el YAML/DSL es válido;
- las reglas normativas llevan documento, apartado, página y cita.

No se puede aceptar una regla que falle algún check.

Las fórmulas internas de las plantillas ST (`=SUM(...)`, `+`/`-`) se convierten en reglas
`plantilla` sin LLM.

## Limitaciones conocidas

- `concepts/st2025.yaml`, `finrep_*.yaml`, `corep_*.yaml` y `rules/st2025/` están **vacíos**
  hasta tener los documentos y la taxonomía EBA.
- `xbrl_csv` y `xbrl_xml` no se han probado con paquetes reales; necesitan un CSV de mapeo
  datapoint → celda sacado del DPM.
- El detector `excel_annex` asume códigos de fila/columna como **texto**. Revisa la pestaña
  «Carga de datos» la primera vez que uses un export nuevo.
