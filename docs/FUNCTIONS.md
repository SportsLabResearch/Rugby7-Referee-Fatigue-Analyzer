# Function inventory

Automatic inventory of functions contained in `rugby7_referee_fatigue_analyzer.py`.

Total functions: **79**

## `tr()`

- **Source line:** 171
- **Arguments:** `texto, idioma`

Sin documentación.

## `traducir_valor()`

- **Source line:** 180
- **Arguments:** `valor, idioma`

Sin documentación.

## `traducir_df()`

- **Source line:** 188
- **Arguments:** `df, idioma`

Sin documentación.

## `etiqueta_idioma()`

- **Source line:** 199
- **Arguments:** `etiqueta, idioma`

Sin documentación.

## `limpiar_texto()`

- **Source line:** 207
- **Arguments:** `x`

Sin documentación.

## `normalizar()`

- **Source line:** 213
- **Arguments:** `x`

Sin documentación.

## `normalizar_simple()`

- **Source line:** 220
- **Arguments:** `x`

Sin documentación.

## `nombre_seguro()`

- **Source line:** 227
- **Arguments:** `x`

Sin documentación.

## `etiqueta_x()`

- **Source line:** 234
- **Arguments:** `txt, ancho`

Sin documentación.

## `detectar_columna()`

- **Source line:** 242
- **Arguments:** `df, candidatos, posicion`

Sin documentación.

## `log_debug()`

- **Source line:** 266
- **Arguments:** `mensaje`

Sin documentación.

## `seleccionar_modo_ejecucion()`

- **Source line:** 271
- **Arguments:** `None`

Sin documentación.

## `seleccionar_nivel_salida()`

- **Source line:** 292
- **Arguments:** `None`

Sin documentación.

## `carpeta_datos()`

- **Source line:** 310
- **Arguments:** `modo`

Sin documentación.

## `seleccionar_hoja_datos()`

- **Source line:** 323
- **Arguments:** `path, solo_nombre`

Selecciona la hoja de datos principal dentro de un Excel único.

## `buscar_excel_principal()`

- **Source line:** 345
- **Arguments:** `datos`

Busca automáticamente el Excel de datos dentro de la carpeta Datos sin depender del nombre. No excluye archivos por contener palabras como datos_arbitros. Si hay varios Excel, prioriza el que contenga columnas/hojas compatibles con la base de partidos-cuestionarios.

## `puntuar_excel()`

- **Source line:** 360
- **Arguments:** `path`

Sin documentación.

## `buscar_plantilla()`

- **Source line:** 399
- **Arguments:** `datos`

Sin documentación.

## `buscar_excel_datos_arbitros()`

- **Source line:** 408
- **Arguments:** `datos`

Sin documentación.

## `leer_excel()`

- **Source line:** 424
- **Arguments:** `path, sheet_name`

Sin documentación.

## `leer_hoja_opcional()`

- **Source line:** 437
- **Arguments:** `path, nombres`

Sin documentación.

## `valores_unicos()`

- **Source line:** 455
- **Arguments:** `df, col`

Sin documentación.

## `es_columna_competitiva()`

- **Source line:** 463
- **Arguments:** `df, col`

Sin documentación.

## `detectar_columnas()`

- **Source line:** 472
- **Arguments:** `df`

Detección adaptada a la base única. Usa la columna específica categoria_competitiva si existe. El análisis por género queda eliminado.

## `seleccionar()`

- **Source line:** 513
- **Arguments:** `titulo, opciones, permitir_todos`

Sin documentación.

## `seleccionar_idioma()`

- **Source line:** 548
- **Arguments:** `None`

Sin documentación.

## `seleccionar_formato()`

- **Source line:** 565
- **Arguments:** `None`

Sin documentación.

## `seleccionar_tipo_analisis()`

- **Source line:** 581
- **Arguments:** `disponibles`

Sin documentación.

## `seleccionar_indices()`

- **Source line:** 603
- **Arguments:** `None`

Sin documentación.

## `preparar_genero()`

- **Source line:** 623
- **Arguments:** `df, col`

Sin documentación.

## `std()`

- **Source line:** 627
- **Arguments:** `x`

Sin documentación.

## `filtrar()`

- **Source line:** 640
- **Arguments:** `df, col, seleccion`

Sin documentación.

## `detectar_variables()`

- **Source line:** 648
- **Arguments:** `df`

Detecta variables PRE y POST cuando existen. Devuelve estructura interna: {"RPE": {"pre": col_pre, "post": col_post}, ...}

## `score_col()`

- **Source line:** 667
- **Arguments:** `nc, patrones, tokens`

Sin documentación.

## `convertir_variables()`

- **Source line:** 703
- **Arguments:** `df, variables`

Sin documentación.

## `preparar_variables_pre_post()`

- **Source line:** 712
- **Arguments:** `df, variables`

Añade columnas de diferencia POST-PRE y genera el diccionario de variables para tablas/gráficos.

## `buscar_fila_arbitro()`

- **Source line:** 743
- **Arguments:** `df_info, arbitro`

Sin documentación.

## `fila_a_diccionario()`

- **Source line:** 759
- **Arguments:** `row`

Sin documentación.

## `buscar_valor_datos()`

- **Source line:** 772
- **Arguments:** `datos_arbitro, etiqueta`

Sin documentación.

## `ruta_foto_desde_datos()`

- **Source line:** 807
- **Arguments:** `datos_dir, datos_arbitro`

Sin documentación.

## `calcular_indices_fatiga()`

- **Source line:** 827
- **Arguments:** `df, variables, indices_seleccionados`

Calcula índices PRE, POST y diferencia POST-PRE cuando existen datos suficientes. IGFA simple PRE/POST = [RPE + MS + (10-PF) + (10-MF) + (10-RC)] / 5 IGFA ponderado PRE/POST = 0.30*RPE + 0.25*MS + 0.20*(10-PF) + 0.15*(10-MF) + 0.10*(10-RC) Delta = POST - PRE

## `get_var()`

- **Source line:** 843
- **Arguments:** `label, momento`

Sin documentación.

## `clasificar_igfa()`

- **Source line:** 883
- **Arguments:** `valor, idioma`

Sin documentación.

## `tabla_partidos()`

- **Source line:** 899
- **Arguments:** `df, col_partido, variables, indices_variables, idioma`

Sin documentación.

## `resumen_por_grupo()`

- **Source line:** 932
- **Arguments:** `df, grupo, variables, nombre_grupo`

Sin documentación.

## `grafico_variable_por_partidos()`

- **Source line:** 949
- **Arguments:** `df, col_partido, etiqueta, col_variable, ruta, idioma`

Sin documentación.

## `grafico_variable_por_grupo()`

- **Source line:** 1007
- **Arguments:** `resumen, columna_grupo, etiqueta, ruta, titulo, idioma`

Sin documentación.

## `lectura_tendencia()`

- **Source line:** 1059
- **Arguments:** `etiqueta, tendencia, idioma`

Sin documentación.

## `interpretar_partido()`

- **Source line:** 1087
- **Arguments:** `etiqueta, stats, idioma`

Sin documentación.

## `interpretar_grupo()`

- **Source line:** 1102
- **Arguments:** `etiqueta, stats, tipo, idioma`

Sin documentación.

## `valor_extremo()`

- **Source line:** 1117
- **Arguments:** `df, col_partido, col, peor_alto`

Sin documentación.

## `tabla_resumen_global()`

- **Source line:** 1127
- **Arguments:** `df, col_partido, variables_analisis, col_competitiva, idioma`

Sin documentación.

## `hallazgos_principales()`

- **Source line:** 1157
- **Arguments:** `df, col_partido, variables_base, variables_analisis, col_competitiva, idioma`

Sin documentación.

## `add()`

- **Source line:** 1159
- **Arguments:** `txt`

Sin documentación.

## `interpretacion_integrada()`

- **Source line:** 1211
- **Arguments:** `hallazgos, idioma`

Sin documentación.

## `conclusion_cientifica_automatica()`

- **Source line:** 1217
- **Arguments:** `hallazgos, idioma`

Sin documentación.

## `set_cell()`

- **Source line:** 1228
- **Arguments:** `cell, text, bold, size`

Sin documentación.

## `add_heading()`

- **Source line:** 1238
- **Arguments:** `doc, text, size`

Sin documentación.

## `add_table_df()`

- **Source line:** 1247
- **Arguments:** `doc, df, idioma`

Sin documentación.

## `tabla_resumen_vertical()`

- **Source line:** 1266
- **Arguments:** `df, columna_grupo`

Cambia orientación de tablas resumen anchas: Métrica | grupo_1 | grupo_2 | grupo_3 ...

## `add_table_resumen_vertical()`

- **Source line:** 1284
- **Arguments:** `doc, df, columna_grupo, idioma`

Sin documentación.

## `tabla_siglas_variables()`

- **Source line:** 1288
- **Arguments:** `idioma`

Sin documentación.

## `insertar_tabla_siglas_y_referencias()`

- **Source line:** 1313
- **Arguments:** `doc, idioma`

Sin documentación.

## `limpiar_num_epigrafe()`

- **Source line:** 1341
- **Arguments:** `txt`

Sin documentación.

## `sentido_variable()`

- **Source line:** 1346
- **Arguments:** `etiqueta`

Criterio conceptual unificado. Variables de carga/fatiga: valores altos = peor estado. Variables positivas: valores altos = mejor estado. Deltas: siempre POST - PRE. - Î”RPE, Î”MS, Î”IGFA > 0 = empeoramiento. - ΔPF, ΔMF, ΔRC > 0 = mejora/disposición más favorable.

## `criterio_delta_texto()`

- **Source line:** 1371
- **Arguments:** `etiqueta, idioma`

Sin documentación.

## `buscar_foto()`

- **Source line:** 1382
- **Arguments:** `datos_dir, arbitro, datos_arbitro`

Búsqueda robusta de fotografía del árbitro en Datos y subcarpetas.

## `insertar_foto_en_documento()`

- **Source line:** 1431
- **Arguments:** `doc, ruta_foto, ancho`

Inserta la foto en la primera celda que contenga 'Foto'. Si no existe, crea un bloque de foto.

## `rellenar_tabla_datos_y_foto()`

- **Source line:** 1462
- **Arguments:** `doc, datos_arbitro, ruta_foto`

Sin documentación.

## `interpretar_partido()`

- **Source line:** 1484
- **Arguments:** `etiqueta, stats, idioma`

Sin documentación.

## `interpretar_grupo()`

- **Source line:** 1510
- **Arguments:** `etiqueta, stats, tipo, idioma`

Sin documentación.

## `tabla_comparativa_dificultad()`

- **Source line:** 1525
- **Arguments:** `df, col_competitiva, variables_analisis, idioma`

Sin documentación.

## `grafico_heatmap_partido_variables()`

- **Source line:** 1547
- **Arguments:** `df, col_partido, variables_analisis, ruta, idioma`

Sin documentación.

## `discutir_partidos_criticos()`

- **Source line:** 1585
- **Arguments:** `df, col_partido, variables_analisis, col_competitiva, idioma`

Sin documentación.

## `discusion_integrada_una_pagina()`

- **Source line:** 1606
- **Arguments:** `df, col_partido, variables_analisis, col_competitiva, hallazgos, idioma`

Sin documentación.

## `crear_informe()`

- **Source line:** 1656
- **Arguments:** `ruta_plantilla, ruta_salida, datos_arbitro, ruta_foto, filtros, tabla_p, resumenes, figuras, tipos, n_obs, resumen_global, hallazgos, interpretacion, conclusion_auto, tabla_dificultad, tabla_criticos, discusion_integrada, idioma`

Sin documentación.

## `heading()`

- **Source line:** 1660
- **Arguments:** `titulo, size`

Sin documentación.

## `main()`

- **Source line:** 1721
- **Arguments:** `None`

Sin documentación.

## `construir_figuras()`

- **Source line:** 1815
- **Arguments:** `idioma_actual`

Sin documentación.
