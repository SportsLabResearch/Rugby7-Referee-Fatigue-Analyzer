from pathlib import Path
import ast

SOURCE = Path("rugby7_referee_fatigue_analyzer.py")
OUTPUT = Path("docs/MODULE_MAP.md")

texto = SOURCE.read_text(encoding="utf-8")
arbol = ast.parse(texto)

categorias = {
    "cli.py": [],
    "constants.py": [],
    "language.py": [],
    "utils.py": [],
    "excel_io.py": [],
    "analysis.py": [],
    "plots.py": [],
    "reports.py": [],
    "referee.py": [],
    "other.py": [],
}

reglas = {
    "cli.py": (
        "seleccionar",
        "main",
    ),
    "language.py": (
        "traduc",
        "idioma",
        "etiqueta_idioma",
        "tr",
    ),
    "utils.py": (
        "limpiar",
        "normalizar",
        "nombre_seguro",
        "etiqueta_x",
        "detectar_columna",
        "valores_unicos",
        "filtrar",
    ),
    "excel_io.py": (
        "excel",
        "hoja",
        "carpeta_datos",
        "buscar_plantilla",
        "tabla_partidos",
    ),
    "analysis.py": (
        "calcular",
        "clasificar",
        "resumen",
        "interpret",
        "hallazgos",
        "conclusion",
        "sentido",
        "criterio",
        "valor_extremo",
        "discutir",
        "discusion",
        "preparar_variables",
        "detectar_variables",
        "convertir_variables",
    ),
    "plots.py": (
        "grafico",
        "heatmap",
    ),
    "reports.py": (
        "crear_informe",
        "add_",
        "set_cell",
        "insertar_",
        "tabla_siglas",
        "tabla_resumen_vertical",
        "rellenar_tabla",
    ),
    "referee.py": (
        "arbitro",
        "foto",
        "fila_a_diccionario",
        "buscar_valor_datos",
    ),
}

funciones = []

for nodo in arbol.body:
    if not isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
        continue

    nombre = nodo.name
    modulo = "other.py"

    for categoria, prefijos in reglas.items():
        if any(prefijo in nombre.lower() for prefijo in prefijos):
            modulo = categoria
            break

    categorias[modulo].append((nombre, nodo.lineno))
    funciones.append((nombre, modulo, nodo.lineno))

lineas = [
    "# Proposed module map",
    "",
    "This document proposes a future modular structure without changing",
    "the current application behaviour.",
    "",
]

for modulo, elementos in categorias.items():
    lineas.append(f"## `{modulo}`")
    lineas.append("")

    if not elementos:
        lineas.append("_No functions assigned._")
        lineas.append("")
        continue

    for nombre, linea in sorted(elementos, key=lambda x: x[1]):
        lineas.append(f"- `{nombre}()` — current line {linea}")

    lineas.append("")

lineas.extend([
    "## Proposed package structure",
    "",
    "```text",
    "rugby7_referee_fatigue_analyzer/",
    "├── __init__.py",
    "├── cli.py",
    "├── constants.py",
    "├── language.py",
    "├── utils.py",
    "├── excel_io.py",
    "├── analysis.py",
    "├── plots.py",
    "├── reports.py",
    "└── referee.py",
    "```",
])

OUTPUT.write_text("\n".join(lineas), encoding="utf-8")

print(f"Mapa creado: {OUTPUT}")
print(f"Funciones clasificadas: {len(funciones)}")

for modulo, elementos in categorias.items():
    print(f"{modulo}: {len(elementos)}")
