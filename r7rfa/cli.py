"""Command-line interface and user-selection functions."""

from r7rfa.utils import normalizar

def seleccionar_modo_ejecucion():
    print("=" * 64)
    print("RUGBY7 REFEREE FATIGUE ANALYZER")
    print("SportsLabResearch")
    print("=" * 64)

    print("\nMODO DE EJECUCIÓN")
    print("\n1. Modo de investigación")
    print("   Utiliza la base de datos privada local.")
    print("\n2. Modo de demostración")
    print("   Utiliza la base de datos pública anonimizada.")

    while True:
        op = input("\nSelección: ").strip()
        if op == "1":
            return "privado"
        if op == "2":
            return "publico"
        print("Selección no válida. Elige 1 o 2.")

def seleccionar(titulo, opciones, permitir_todos=True):
    if not opciones:
        return []
    print(f"\n{titulo}")
    for i, op in enumerate(opciones, 1):
        print(f"{i}. {op}")
    if permitir_todos:
        print(f"{len(opciones)+1}. Todos/as")

    while True:
        txt = input("Selecciona opción/es (ej.: 1 | 1,3 | TODOS): ").strip()
        if permitir_todos and normalizar(txt) in ["todos", "todas", "all", str(len(opciones)+1)]:
            return opciones
        partes = [p.strip() for p in txt.split(",") if p.strip()]
        idx = []
        ok = True
        for p in partes:
            if not p.isdigit():
                ok = False
                break
            n = int(p)
            if 1 <= n <= len(opciones):
                idx.append(n-1)
            elif permitir_todos and n == len(opciones)+1:
                return opciones
            else:
                ok = False
                break
        if ok and idx:
            return [opciones[i] for i in idx]
        print("Selección no válida.")

def seleccionar_idioma():
    print("\nIDIOMA DEL INFORME")
    print("1. Inglés")
    print("2. Castellano")
    print("3. Ambos")

    while True:
        op = input("Selecciona idioma: ").strip()
        if op == "1":
            return ["ingles"]
        elif op == "2":
            return ["castellano"]
        elif op == "3":
            return ["ingles", "castellano"]
        print("Opción no válida.")

def seleccionar_formato():
    print("\nFORMATO DE SALIDA")
    print("1. Word (.docx)")
    print("2. PDF (.pdf)")
    print("3. Ambos")

    while True:
        op = input("Selecciona formato: ").strip()
        if op == "1":
            return ["docx"]
        elif op == "2":
            return ["pdf"]
        elif op == "3":
            return ["docx", "pdf"]
        print("Opción no válida.")

def seleccionar_tipo_analisis(disponibles):
    print("\nTIPO DE ANÃLISIS")
    ops = []
    if disponibles.get("partidos"):
        ops.append(("partidos", "Por partidos"))
    if disponibles.get("competitiva"):
        ops.append(("competitiva", "Por categoría competitiva"))
    ops.append(("todos", "Todos"))

    for i, (_, txt) in enumerate(ops, 1):
        print(f"{i}. {txt}")

    while True:
        op = input("Selecciona tipo de análisis: ").strip()
        if op.isdigit() and 1 <= int(op) <= len(ops):
            key = ops[int(op)-1][0]
            if key == "todos":
                return {k for k, v in disponibles.items() if v and k != "genero"}
            return {key}
        print("Opción no válida.")

def seleccionar_indices():
    print("\nÃNDICES DE FATIGA A INCLUIR")
    print("1. IGFA simple: media de RPE, MS y variables positivas invertidas")
    print("2. IGFA ponderado: mayor peso para RPE y molestias")
    print("3. Ambos índices")
    print("4. No incluir índices")

    while True:
        op = input("Selecciona índice/s: ").strip()
        if op == "1":
            return {"IGFA simple"}
        elif op == "2":
            return {"IGFA ponderado"}
        elif op == "3":
            return {"IGFA simple", "IGFA ponderado"}
        elif op == "4":
            return set()
        print("Opción no válida.")
