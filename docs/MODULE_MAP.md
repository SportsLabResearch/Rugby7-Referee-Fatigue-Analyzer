# Proposed module map

This document proposes a future modular structure without changing
the current application behaviour.

## `cli.py`

- `seleccionar_modo_ejecucion()` — current line 271
- `seleccionar_nivel_salida()` — current line 292
- `seleccionar_hoja_datos()` — current line 323
- `seleccionar()` — current line 513
- `seleccionar_idioma()` — current line 548
- `seleccionar_formato()` — current line 565
- `seleccionar_tipo_analisis()` — current line 581
- `seleccionar_indices()` — current line 603
- `main()` — current line 1721

## `constants.py`

_No functions assigned._

## `language.py`

- `tr()` — current line 171
- `traducir_valor()` — current line 180
- `traducir_df()` — current line 188
- `etiqueta_idioma()` — current line 199
- `buscar_excel_datos_arbitros()` — current line 408
- `filtrar()` — current line 640
- `buscar_fila_arbitro()` — current line 743
- `valor_extremo()` — current line 1117

## `utils.py`

- `limpiar_texto()` — current line 207
- `normalizar()` — current line 213
- `normalizar_simple()` — current line 220
- `nombre_seguro()` — current line 227
- `etiqueta_x()` — current line 234
- `detectar_columna()` — current line 242
- `valores_unicos()` — current line 455
- `detectar_columnas()` — current line 472
- `limpiar_num_epigrafe()` — current line 1341

## `excel_io.py`

- `carpeta_datos()` — current line 310
- `buscar_excel_principal()` — current line 345
- `buscar_plantilla()` — current line 399
- `leer_excel()` — current line 424
- `leer_hoja_opcional()` — current line 437
- `tabla_partidos()` — current line 899

## `analysis.py`

- `detectar_variables()` — current line 648
- `convertir_variables()` — current line 703
- `preparar_variables_pre_post()` — current line 712
- `calcular_indices_fatiga()` — current line 827
- `clasificar_igfa()` — current line 883
- `resumen_por_grupo()` — current line 932
- `interpretar_partido()` — current line 1087
- `interpretar_grupo()` — current line 1102
- `tabla_resumen_global()` — current line 1127
- `hallazgos_principales()` — current line 1157
- `interpretacion_integrada()` — current line 1211
- `conclusion_cientifica_automatica()` — current line 1217
- `tabla_resumen_vertical()` — current line 1266
- `add_table_resumen_vertical()` — current line 1284
- `sentido_variable()` — current line 1346
- `criterio_delta_texto()` — current line 1371
- `interpretar_partido()` — current line 1484
- `interpretar_grupo()` — current line 1510
- `discutir_partidos_criticos()` — current line 1585
- `discusion_integrada_una_pagina()` — current line 1606

## `plots.py`

- `grafico_variable_por_partidos()` — current line 949
- `grafico_variable_por_grupo()` — current line 1007
- `grafico_heatmap_partido_variables()` — current line 1547

## `reports.py`

- `set_cell()` — current line 1228
- `add_heading()` — current line 1238
- `add_table_df()` — current line 1247
- `tabla_siglas_variables()` — current line 1288
- `insertar_tabla_siglas_y_referencias()` — current line 1313
- `insertar_foto_en_documento()` — current line 1431
- `rellenar_tabla_datos_y_foto()` — current line 1462
- `crear_informe()` — current line 1656

## `referee.py`

- `fila_a_diccionario()` — current line 759
- `buscar_valor_datos()` — current line 772
- `ruta_foto_desde_datos()` — current line 807
- `buscar_foto()` — current line 1382

## `other.py`

- `log_debug()` — current line 266
- `es_columna_competitiva()` — current line 463
- `preparar_genero()` — current line 623
- `lectura_tendencia()` — current line 1059
- `tabla_comparativa_dificultad()` — current line 1525

## Proposed package structure

```text
rugby7_referee_fatigue_analyzer/
├── __init__.py
├── cli.py
├── constants.py
├── language.py
├── utils.py
├── excel_io.py
├── analysis.py
├── plots.py
├── reports.py
└── referee.py
```