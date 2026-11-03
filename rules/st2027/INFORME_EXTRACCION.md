# Extracción desde plantillas ST: 2027 EU-wide stress test - Draft Templates.xlsx

Fecha: 2026-11-03

- Reglas de mapeo a COREP: **2489** (294 son alternativas C 09.02 / C 09.05)
- Controles internos (fórmulas) en punto de partida: 21886 (exportados como .opcional)
- Celdas de punto de partida inventariadas: 226352

## Reglas de mapeo por hoja

| Hoja | Reglas |
|---|---|
| CSV_CAP | 61 |
| CSV_CR_REA_IRB_DET | 1827 |
| CSV_CR_REA_OF | 96 |
| CSV_CR_REA_STA_DET | 504 |
| CSV_REA_SUM | 1 |

## Textos de mapeo no convertidos

| Tipo | Nº | Ejemplos |
|---|---|---|
| calculo_interno | 9 | CSV_CAP!S70: [A.1-A.1.13-A.1.22+MIN(A.2+A.1.13-A.2.2-A.2.4+MIN(A.4+A.2.2-A.4.3,0),0)]<br>CSV_CAP!S71: [A.1-A.1.22+A.2-A.2.2-A.2.4+MIN(A.4+A.2.2-A.4.3,0)]<br>CSV_CAP!S72: [A.1-A.1.22+A.2-A.2.4+A.4-A.4.3+A.5)] |
| columna_compuesta_no_soportada | 3 | CSV_CR_REA_IRB_DET!W8: (C 09.05, c0125 or C 09.02, c0125) minus C 09.05, c0126<br>CSV_CR_REA_IRB_DET!Y8: (C 09.05, c0130 or C 09.02, c0130) minus C 09.05, c0131<br>CSV_CR_REA_IRB_DET!AA8: (C 09.05, c0050+c0055 or C 09.02, c0050+c0055) minus (C 09.05, c0051+c0056) |
| porcentaje | 5 | CSV_CAP!S105: C 03.00 (r0130,c0010) % <br>CSV_CAP!S106: C 03.00 (r0140,c0010) % <br>CSV_CAP!S107: C 03.00 (r0160,c0010) %  |
| proyeccion | 1548 | CSV_CR_REA_STA_DET!P178:  projection(C 09.01, r0010+r0020+r0030+r0040+r0050+r0060)<br>CSV_CR_REA_STA_DET!P179:  projection(C 09.01, r0070)<br>CSV_CR_REA_STA_DET!P180:  projection(C 09.01, r0080) |

## Hojas

| Hoja | Filas | Columnas | Filas punto partida | Cols punto partida | Avisos |
|---|---|---|---|---|---|
| CSV_CR_SUM | 24 | 7 | 24 | 1 |  |
| CSV_CR_SCEN | 2112 | 54 | 528 | 54 |  |
| CSV_CR_SCEN_OFF_BS | 154 | 18 | 22 | 18 |  |
| CSV_CR_SECTOR | 2208 | 46 | 552 | 46 |  |
| CSV_CR_NPL | 48 | 77 | 48 | 11 |  |
| CSV_CR_SEC_SUM | 24 | 1 | 24 | 1 |  |
| CSV_CR_SEC | 28 | 62 | 4 | 62 |  |
| CSV_CR_REA_STA_DET | 420 | 3 | 168 | 3 |  |
| CSV_CR_REA_IRB_DET | 1470 | 15 | 210 | 15 |  |
| CSV_CR_REA_OF | 60 | 4 | 24 | 4 |  |
| CSV_CR_REA_STA | 29 | 21 | 29 | 3 |  |
| CSV_CR_REA_IRB | 38 | 21 | 38 | 3 |  |
| CSV_MR_SUM | 43 | 8 | 43 | 8 |  |
| CSV_MR_PROJ | 23 | 11 | 23 | 11 |  |
| CSV_MR_CCR | 47 | 17 | 47 | 10 |  |
| CSV_MR_REA | 45 | 8 | 45 | 1 |  |
| CSV_NII_SUM | 88 | 7 | 88 | 1 |  |
| CSV_NII_CALC | 943 | 271 | 943 | 104 |  |
| CSV_NII_CALC_FUNDING_MATCH | 14 | 271 | 14 | 104 |  |
| CSV_OR_GEN | 25 | 13 | 25 | 13 |  |
| CSV_OR_CON | 17 | 17 | 17 | 17 |  |
| CSV_REA_SUM | 38 | 7 | 38 | 1 |  |
| CSV_NFCI_DIV | 7 | 7 | 7 | 1 |  |
| CSV_ADM_EXP | 70 | 7 | 70 | 1 |  |
| CSV_ONEOFF | 10 | 21 | 10 | 21 |  |
| CSV_MDA | 35 | 10 | 35 | 1 |  |
| CSV_CAP | 104 | 7 | 104 | 1 |  |
| CSV_CAPMEAS | 7 | 1 | 7 | 1 |  |
| CSV_CL_SUM | 20 | 6 | 20 | 2 |  |
| CSV_CL_TR | 3480 | 47 | 1392 | 47 |  |
