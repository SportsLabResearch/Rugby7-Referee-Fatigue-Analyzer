from pathlib import Path
import ast

SOURCE = Path("rugby7_referee_fatigue_analyzer.py")
OUTPUT = Path("docs/FUNCTIONS.md")

texto = SOURCE.read_text(encoding="utf-8")
arbol = ast.parse(texto)

funciones = []

for nodo in ast.walk(arbol):
    if not isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
        continue

    argumentos = []

    for arg in nodo.args.args:
        argumentos.append(arg.arg)

    if nodo.args.vararg:
        argumentos.append(f"*{nodo.args.vararg.arg}")

    for arg in nodo.args.kwonlyargs:
        argumentos.append(arg.arg)

    if nodo.args.kwarg:
        argumentos.append(f"**{nodo.args.kwarg.arg}")

    docstring = ast.get_docstring(nodo) or "Sin documentación."
    docstring = " ".join(docstring.split())

    funciones.append({
        "nombre": nodo.name,
        "linea": nodo.lineno,
        "argumentos": ", ".join(argumentos),
        "documentacion": docstring,
    })

funciones.sort(key=lambda item: item["linea"])

lineas = [
    "# Function inventory",
    "",
    "Automatic inventory of functions contained in "
    "`rugby7_referee_fatigue_analyzer.py`.",
    "",
    f"Total functions: **{len(funciones)}**",
    "",
]

for funcion in funciones:
    lineas.extend([
        f"## `{funcion['nombre']}()`",
        "",
        f"- **Source line:** {funcion['linea']}",
        f"- **Arguments:** `{funcion['argumentos'] or 'None'}`",
        "",
        funcion["documentacion"],
        "",
    ])

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
OUTPUT.write_text("\n".join(lineas), encoding="utf-8")

print(f"Inventario creado: {OUTPUT}")
print(f"Funciones detectadas: {len(funciones)}")
