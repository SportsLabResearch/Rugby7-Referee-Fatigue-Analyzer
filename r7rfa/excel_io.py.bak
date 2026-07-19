"""Excel, template and data-directory input utilities."""

import sys
from pathlib import Path

import pandas as pd

from r7rfa.utils import limpiar_texto, normalizar

def carpeta_datos(modo="privado"):
    if modo == "publico":
        d = Path.cwd() / "examples" / "anonymized_data"
    else:
        d = Path.cwd() / "Datos"

    if not d.exists():
        print(f"ERROR: no existe la carpeta de datos: {d}")
        sys.exit(1)

    return d

def seleccionar_hoja_datos(path, solo_nombre=False):
    """Selecciona la hoja de datos principal dentro de un Excel único."""
    xls = pd.ExcelFile(path)
    preferidas = [
        "partidos_cuestionarios",
        "partidos_nombre_arbitro",
        "partidos",
        "cuestionarios",
    ]
    mapa = {normalizar(h): h for h in xls.sheet_names}
    for h in preferidas:
        if h in mapa:
            return mapa[h]

    # Si no hay nombre preferido, se queda con la primera hoja que no parezca diccionario/preguntas.
    for h in xls.sheet_names:
        nh = normalizar(h)
        if nh not in ["preguntas", "diccionario", "diccionario_variables"]:
            return h
    return xls.sheet_names[0]

def buscar_excel_principal(datos):
    """
    Busca automáticamente el Excel de datos dentro de la carpeta Datos sin depender del nombre.
    No excluye archivos por contener palabras como datos_arbitros.
    Si hay varios Excel, prioriza el que contenga columnas/hojas compatibles con la base de partidos-cuestionarios.
    """
    excels = []
    for ext in ("*.xlsx", "*.xls", "*.xlsm"):
        excels.extend(datos.glob(ext))
    excels = [p for p in excels if not p.name.startswith("~$")]

    if not excels:
        print("ERROR: no se encontró ningún archivo Excel en la carpeta Datos.")
        sys.exit(1)

    def puntuar_excel(path):
        score = 0
        try:
            xls = pd.ExcelFile(path)
            hojas_norm = [normalizar(h) for h in xls.sheet_names]
            for h in hojas_norm:
                if h in ["partidos_cuestionarios", "partidos_nombre_arbitro", "partidos", "cuestionarios"]:
                    score += 20
            hoja = seleccionar_hoja_datos(path, solo_nombre=True)
            df0 = pd.read_excel(path, sheet_name=hoja, nrows=40)
            cols = [normalizar(c) for c in df0.columns]
            claves = [
                "match_id", "partido", "ref_id", "referee", "arbitro",
                "rpe", "rpe_pre", "rpe_post", "ms_pre", "ms_post",
                "pf_pre", "pf_post", "mf_pre", "mf_post", "rc_pre", "rc_post"
            ]
            for c in cols:
                for k in claves:
                    if k in c or c in k:
                        score += 3
            # Evita elegir Excel claramente vacíos o diccionarios si hay otro mejor.
            score += min(len(cols), 30) / 10
        except Exception:
            score = -1
        return score

    candidatos = sorted(excels, key=lambda p: (puntuar_excel(p), p.stat().st_mtime), reverse=True)
    elegido = candidatos[0]

    if puntuar_excel(elegido) < 0:
        print("ERROR: se encontraron Excel en Datos, pero no se pudieron leer.")
        for p in excels:
            print(f"- {p.name}")
        sys.exit(1)

    log_debug(f"Excel principal detectado: {elegido.name}")
    return elegido

def buscar_plantilla(datos):
    docs = [p for p in datos.glob("*.docx") if "plantilla" in normalizar(p.stem) and not p.name.startswith("~$")]
    if not docs:
        print("ERROR: no se encontró plantilla.docx en Datos.")
        sys.exit(1)
    log_debug(f"Plantilla detectada: {docs[0].name}")
    return docs[0]

def buscar_excel_datos_arbitros(datos):
    excels = []
    for ext in ("*.xlsx", "*.xls", "*.xlsm"):
        excels.extend(datos.glob(ext))
    docs = [
        p for p in excels
        if not p.name.startswith("~$")
        and ("datos_arbitro" in normalizar(p.stem) or "datos_arbitros" in normalizar(p.stem))
    ]
    if not docs:
        print("AVISO: no se encontró datos_arbitros.xlsx.")
        return None
    print(f"Datos de árbitros detectado: {docs[0].name}")
    return docs[0]

def leer_excel(path, sheet_name=None):
    if sheet_name is None:
        sheet_name = seleccionar_hoja_datos(path)
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = df.dropna(how="all").copy()
    df.columns = [limpiar_texto(c) for c in df.columns]
    if df.empty:
        print(f"ERROR: {path.name} / hoja {sheet_name} está vacío.")
        sys.exit(1)
    log_debug(f"Hoja de datos utilizada: {sheet_name}")
    return df

def leer_hoja_opcional(path, nombres):
    try:
        xls = pd.ExcelFile(path)
        mapa = {normalizar(h): h for h in xls.sheet_names}
        for n in nombres:
            hn = normalizar(n)
            if hn in mapa:
                df = pd.read_excel(path, sheet_name=mapa[hn])
                df = df.dropna(how="all").copy()
                df.columns = [limpiar_texto(c) for c in df.columns]
                if not df.empty:
                    log_debug(f"Hoja auxiliar utilizada: {mapa[hn]}")
                    return df
    except Exception:
        pass
    return None
