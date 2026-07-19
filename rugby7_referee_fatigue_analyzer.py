# -*- coding: utf-8 -*-
"""
SCRIPT DE INFORME AUTOMÃTICO PARA ÃRBITROS DE RUGBY 7
V30 - informe profesional Q1 + foto robusta + dificultad competitiva + heatmap

Estructura esperada:
    Carpeta_del_script/
    â”œâ”€â”€ script_informe_arbitro_rugby7_V16_PLANTILLA_TABLA_FOTO_PARTIDOS.py
    â””â”€â”€ Datos/
        â”œâ”€â”€ cuestionarios.xlsx
        â”œâ”€â”€ plantilla.docx
        â”œâ”€â”€ datos_arbitros.xlsx
        â”œâ”€â”€ Raul.jpg / fotos/Raul.png / imagenes/Raul.jpeg
        â””â”€â”€ ...

Correcciones principales:
1. NO crea una tabla nueva para los datos del Ã¡rbitro si la plantilla ya tiene la tabla Variable / Valor / Foto.
2. Rellena la columna "Valor" de esa tabla usando datos_arbitros.xlsx.
3. Inserta la foto en la celda grande donde aparece "Foto".
4. Si se selecciona anÃ¡lisis por partidos:
   - Genera un grÃ¡fico independiente por cada variable.
   - Cada grÃ¡fico muestra todos los partidos seleccionados.
   - Cada punto es el valor directo del partido, no media.
   - No calcula SD/min/max por partido porque n=1.
5. AÃ±ade el anÃ¡lisis debajo de la ficha inicial de la plantilla.
"""

import re
import sys
import textwrap
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from docx2pdf import convert as docx2pdf_convert
    PDF_OK = True
except Exception:
    PDF_OK = False

try:
    import matplotlib.pyplot as plt
except Exception:
    print("ERROR: falta matplotlib. Instala con: pip install matplotlib")
    sys.exit(1)

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
except Exception:
    print("ERROR: falta python-docx. Instala con: pip install python-docx")
    sys.exit(1)


VARIABLES_OBJETIVO = {
    "RPE": ["rpe", "percepcion", "percepciÃ³n", "rendimiento", "esfuerzo"],
    "Molestias (MS)": ["molestia", "molestias", "ms"],
    "Capacidad fÃ­sica (PF)": ["capacidad fisica", "capacidad fÃ­sica", "pf"],
    "Capacidad mental (MF)": ["capacidad mental", "mf"],
    "Confianza (RC)": ["confianza", "rc"],
}

VALORES_COMPETITIVOS = ["Desequilibrado", "Moderadamente equilibrado", "Muy equilibrado"]
EXTENSIONES_IMAGEN = [".jpg", ".jpeg", ".png", ".bmp", ".webp"]
COLOR_TITULO = RGBColor(31, 78, 121)

# =============================================================================
# V31: sin numeraciÃ³n en epÃ­grafes, hallazgos en viÃ±etas y tendencia+ecuaciÃ³n en todas las grÃ¡ficas.
# =============================================================================

# =============================================================================
# IDIOMA / TRADUCCIÃ“N INTERNA DEL INFORME
# =============================================================================

TEXTOS = {
    "castellano": {
        "sin_datos": "Sin datos.",
        "informe": "Informe de anÃ¡lisis",
        "observaciones": "Observaciones analizadas",
        "tipo_analisis_incluido": "Tipo de anÃ¡lisis incluido",
        "indices_incluidos": "Ãndices de fatiga incluidos",
        "filtros": "Filtros aplicados",
        "partidos": "AnÃ¡lisis por partidos",
        "partidos_intro": "Este anÃ¡lisis presenta valores directos de cada partido. No se calculan desviaciones estÃ¡ndar por partido porque cada partido tiene un Ãºnico registro.",
        "genero": "AnÃ¡lisis por gÃ©nero",
        "competitiva": "AnÃ¡lisis por categorÃ­a competitiva",
        "graficos": "GrÃ¡ficos e interpretaciÃ³n",
        "interpretacion": "InterpretaciÃ³n",
        "conclusion": "ConclusiÃ³n",
        "conclusion_partidos": "El anÃ¡lisis por partidos permite revisar la evoluciÃ³n real del Ã¡rbitro partido a partido. Cada grÃ¡fico representa una variable concreta en todos los partidos seleccionados.",
        "conclusion_genero": "El anÃ¡lisis por gÃ©nero resume las diferencias agrupadas por M/W cuando existen datos suficientes.",
        "conclusion_competitiva": "El anÃ¡lisis por categorÃ­a competitiva resume la respuesta del Ã¡rbitro segÃºn el equilibrio del partido.",
        "variables": "DescripciÃ³n de variables analizadas",
        "variables_intro": "La siguiente tabla resume las siglas empleadas en el informe, su significado y una interpretaciÃ³n breve apoyada en literatura cientÃ­fica sobre monitorizaciÃ³n subjetiva, carga interna y rendimiento.",
        "referencias": "Referencias cientÃ­ficas",
        "puntuacion_directa": "PuntuaciÃ³n directa (0-10)",
        "media_010": "Media (0-10)",
        "por_partido": "por partido",
        "por_partidos": "por partidos",
        "por_genero": "por gÃ©nero",
        "por_competitiva": "por categorÃ­a competitiva",
    },
    "ingles": {
        "sin_datos": "No data.",
        "informe": "Analysis report",
        "observaciones": "Analysed observations",
        "tipo_analisis_incluido": "Type of analysis included",
        "indices_incluidos": "Fatigue indices included",
        "filtros": "Applied filters",
        "partidos": "Match-by-match analysis",
        "partidos_intro": "This analysis presents the direct values recorded for each match. Standard deviations are not calculated per match because each match has a single record.",
        "genero": "Sex-based analysis",
        "competitiva": "Competitive balance category analysis",
        "graficos": "Graphs and interpretation",
        "interpretacion": "Interpretation",
        "conclusion": "Conclusion",
        "conclusion_partidos": "The match-by-match analysis makes it possible to review the referee's actual evolution across the selected matches. Each graph represents one specific variable across all selected matches.",
        "conclusion_genero": "The sex-based analysis summarises grouped differences by M/W when sufficient data are available.",
        "conclusion_competitiva": "The competitive balance category analysis summarises the referee's response according to match balance.",
        "variables": "Description of the analysed variables",
        "variables_intro": "The following table summarises the abbreviations used in the report, their meaning and a brief interpretation supported by scientific literature on subjective monitoring, internal load and performance.",
        "referencias": "Scientific references",
        "puntuacion_directa": "Direct score (0-10)",
        "media_010": "Mean (0-10)",
        "por_partido": "by match",
        "por_partidos": "by match",
        "por_genero": "by sex",
        "por_competitiva": "by competitive balance category",
    },
}

TRAD_EN = {
    "RPE": "RPE",
    "Molestias (MS)": "Muscle soreness (MS)",
    "Capacidad fÃ­sica (PF)": "Physical fitness (PF)",
    "Capacidad mental (MF)": "Mental fitness (MF)",
    "Confianza (RC)": "Referee confidence (RC)",
    "IGFA simple": "Simple GRFI",
    "IGFA ponderado": "Weighted GRFI",
    "ClasificaciÃ³n IGFA simple": "Simple GRFI classification",
    "ClasificaciÃ³n IGFA ponderado": "Weighted GRFI classification",
    "Partido": "Match",
    "GÃ©nero": "Sex",
    "CategorÃ­a competitiva": "Competitive balance category",
    "Resumen por gÃ©nero": "Summary by sex",
    "Resumen por categorÃ­a competitiva": "Summary by competitive balance category",
    "Fatiga muy baja": "Very low fatigue",
    "Fatiga baja": "Low fatigue",
    "Fatiga moderada": "Moderate fatigue",
    "Fatiga alta": "High fatigue",
    "Fatiga muy alta": "Very high fatigue",
    "Sin dato": "No data",
    "No incluidos": "Not included",
    "No aplicable": "Not applicable",
    "Todas/no aplica": "All/not applicable",
    "Todos": "All",
    "Todas": "All",
    "Tipo de anÃ¡lisis": "Type of analysis",
    "Ãndices incluidos": "Included indices",
    "Plantilla": "Template",
    "partidos": "matches",
    "genero": "sex",
    "competitiva": "competitive balance",
}


def tr(texto, idioma="castellano"):
    if idioma != "ingles":
        return texto
    if texto is None:
        return texto
    s = str(texto)
    return TRAD_EN.get(s, s)


def traducir_valor(valor, idioma="castellano"):
    if idioma != "ingles":
        return valor
    if isinstance(valor, list):
        return [traducir_valor(v, idioma) for v in valor]
    return tr(valor, idioma)


def traducir_df(df, idioma="castellano"):
    if idioma != "ingles" or df is None or df.empty:
        return df
    out = df.copy()
    out.columns = [tr(c, idioma) for c in out.columns]
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].map(lambda x: tr(x, idioma) if isinstance(x, str) else x)
    return out


def etiqueta_idioma(etiqueta, idioma="castellano"):
    return tr(etiqueta, idioma)


# =============================================================================
# TEXTO Y DETECCIÃ“N
# =============================================================================

def limpiar_texto(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


def normalizar(x):
    x = limpiar_texto(x).lower()
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    x = re.sub(r"[^a-z0-9]+", "_", x)
    return re.sub(r"_+", "_", x).strip("_")


def normalizar_simple(x):
    x = limpiar_texto(x).lower()
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    x = re.sub(r"[^a-z0-9]+", "", x)
    return x


def nombre_seguro(x):
    x = limpiar_texto(x)
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    x = re.sub(r"[^A-Za-z0-9_-]+", "_", x)
    return re.sub(r"_+", "_", x).strip("_") or "sin_nombre"


def etiqueta_x(txt, ancho=16):
    txt = limpiar_texto(txt)
    txt = txt.replace(" vs ", "\nvs\n").replace(" VS ", "\nvs\n")
    if len(txt) <= ancho:
        return txt
    return "\n".join(textwrap.wrap(txt, width=ancho))


def detectar_columna(df, candidatos, posicion=None):
    cols = list(df.columns)
    norm = {c: normalizar(c) for c in cols}

    for cand in candidatos:
        ncand = normalizar(cand)
        for c, n in norm.items():
            if n == ncand:
                return c

    for cand in candidatos:
        ncand = normalizar(cand)
        for c, n in norm.items():
            if ncand and (ncand in n or n in ncand):
                return c

    if posicion is not None and posicion < len(cols):
        return cols[posicion]
    return None


def seleccionar_modo_ejecucion():
    print("\nMODO DE EJECUCIÓN")
    print("1. Privado/local")
    print("2. Público/anonimizado")

    while True:
        op = input("Selecciona modo: ").strip()
        if op == "1":
            return "privado"
        if op == "2":
            return "publico"
        print("Opción no válida.")


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
    """Selecciona la hoja de datos principal dentro de un Excel Ãºnico."""
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
    Busca automÃ¡ticamente el Excel de datos dentro de la carpeta Datos sin depender del nombre.
    No excluye archivos por contener palabras como datos_arbitros.
    Si hay varios Excel, prioriza el que contenga columnas/hojas compatibles con la base de partidos-cuestionarios.
    """
    excels = []
    for ext in ("*.xlsx", "*.xls", "*.xlsm"):
        excels.extend(datos.glob(ext))
    excels = [p for p in excels if not p.name.startswith("~$")]

    if not excels:
        print("ERROR: no se encontrÃ³ ningÃºn archivo Excel en la carpeta Datos.")
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
            # Evita elegir Excel claramente vacÃ­os o diccionarios si hay otro mejor.
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

    print(f"Excel principal detectado: {elegido.name}")
    return elegido


def buscar_plantilla(datos):
    docs = [p for p in datos.glob("*.docx") if "plantilla" in normalizar(p.stem) and not p.name.startswith("~$")]
    if not docs:
        print("ERROR: no se encontrÃ³ plantilla.docx en Datos.")
        sys.exit(1)
    print(f"Plantilla detectada: {docs[0].name}")
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
        print("AVISO: no se encontrÃ³ datos_arbitros.xlsx.")
        return None
    print(f"Datos de Ã¡rbitros detectado: {docs[0].name}")
    return docs[0]


def leer_excel(path, sheet_name=None):
    if sheet_name is None:
        sheet_name = seleccionar_hoja_datos(path)
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = df.dropna(how="all").copy()
    df.columns = [limpiar_texto(c) for c in df.columns]
    if df.empty:
        print(f"ERROR: {path.name} / hoja {sheet_name} estÃ¡ vacÃ­o.")
        sys.exit(1)
    print(f"Hoja de datos utilizada: {sheet_name}")
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
                    print(f"Hoja auxiliar utilizada: {mapa[hn]}")
                    return df
    except Exception:
        pass
    return None


def valores_unicos(df, col):
    if col is None or col not in df.columns:
        return []
    vals = [limpiar_texto(v) for v in df[col].dropna().unique()]
    vals = [v for v in vals if v]
    return sorted(vals, key=lambda x: normalizar(x))


def es_columna_competitiva(df, col):
    if col is None:
        return False
    vals = {normalizar(v) for v in valores_unicos(df, col)}
    comp = {normalizar(v) for v in VALORES_COMPETITIVOS}
    return bool(vals) and vals.issubset(comp)



def detectar_columnas(df):
    """
    DetecciÃ³n adaptada a la base Ãºnica.
    Usa la columna especÃ­fica categoria_competitiva si existe.
    El anÃ¡lisis por gÃ©nero queda eliminado.
    """
    col_arbitro = detectar_columna(df, ["ref_id", "referee_name", "arbitro", "Ã¡rbitro", "nombre", "referee"], posicion=0)
    col_partido = detectar_columna(df, ["match_id", "partido", "match", "game", "id_partido"], posicion=0)
    col_cat = detectar_columna(df, ["categoria", "categorÃ­a", "nivel", "division", "divisiÃ³n", "category"], posicion=None)

    col_comp = detectar_columna(df, [
        "categoria_competitiva",
        "categorÃ­a_competitiva",
        "categorÃ­a competitiva",
        "dificultad_partido",
        "nivel_competitivo",
        "equilibrio competitivo",
        "competitive_balance"
    ], posicion=None)

    if col_comp is None:
        col_comp = detectar_columna(df, ["competitiva", "indice competitivo", "Ã­ndice competitivo"], posicion=None)

    if col_partido is None and len(df.columns) >= 1:
        col_partido = df.columns[0]

    print("\nCOLUMNAS DETECTADAS")
    print(f"- Ãrbitro: {col_arbitro}")
    print("- GÃ©nero: anÃ¡lisis eliminado")
    print(f"- Partido: {col_partido if col_partido else 'No detectada'}")
    print(f"- CategorÃ­a competitiva: {col_comp if col_comp else 'No detectada'}")

    if col_arbitro is None:
        print("ERROR: no se detectÃ³ la columna del Ã¡rbitro.")
        sys.exit(1)
    if col_partido is None:
        print("ERROR: no se detectÃ³ la columna de partido.")
        sys.exit(1)

    return {"arbitro": col_arbitro, "genero": None, "partido": col_partido, "categoria": col_cat, "competitiva": col_comp}

def seleccionar(titulo, opciones, permitir_todos=True):
    if not opciones:
        return []
    print(f"\n{titulo}")
    for i, op in enumerate(opciones, 1):
        print(f"{i}. {op}")
    if permitir_todos:
        print(f"{len(opciones)+1}. Todos/as")

    while True:
        txt = input("Selecciona opciÃ³n/es (ej.: 1 | 1,3 | TODOS): ").strip()
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
        print("SelecciÃ³n no vÃ¡lida.")




def seleccionar_idioma():
    print("\nIDIOMA DEL INFORME")
    print("1. InglÃ©s")
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
        print("OpciÃ³n no vÃ¡lida.")


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
        print("OpciÃ³n no vÃ¡lida.")

def seleccionar_tipo_analisis(disponibles):
    print("\nTIPO DE ANÃLISIS")
    ops = []
    if disponibles.get("partidos"):
        ops.append(("partidos", "Por partidos"))
    if disponibles.get("competitiva"):
        ops.append(("competitiva", "Por categorÃ­a competitiva"))
    ops.append(("todos", "Todos"))

    for i, (_, txt) in enumerate(ops, 1):
        print(f"{i}. {txt}")

    while True:
        op = input("Selecciona tipo de anÃ¡lisis: ").strip()
        if op.isdigit() and 1 <= int(op) <= len(ops):
            key = ops[int(op)-1][0]
            if key == "todos":
                return {k for k, v in disponibles.items() if v and k != "genero"}
            return {key}
        print("OpciÃ³n no vÃ¡lida.")


def seleccionar_indices():
    print("\nÃNDICES DE FATIGA A INCLUIR")
    print("1. IGFA simple: media de RPE, MS y variables positivas invertidas")
    print("2. IGFA ponderado: mayor peso para RPE y molestias")
    print("3. Ambos Ã­ndices")
    print("4. No incluir Ã­ndices")

    while True:
        op = input("Selecciona Ã­ndice/s: ").strip()
        if op == "1":
            return {"IGFA simple"}
        elif op == "2":
            return {"IGFA ponderado"}
        elif op == "3":
            return {"IGFA simple", "IGFA ponderado"}
        elif op == "4":
            return set()
        print("OpciÃ³n no vÃ¡lida.")


def preparar_genero(df, col):
    if col is None:
        return df

    def std(x):
        n = normalizar(x)
        if n in ["m", "masculino", "male", "hombre", "men"]:
            return "M"
        if n in ["w", "f", "femenino", "female", "mujer", "women"]:
            return "W"
        return limpiar_texto(x)

    df = df.copy()
    df["_GENERO_STD"] = df[col].map(std)
    return df


def filtrar(df, col, seleccion):
    if col is None or not seleccion:
        return df.copy()
    sel = {normalizar(v) for v in seleccion}
    return df[df[col].map(lambda x: normalizar(x) in sel)].copy()



def detectar_variables(df):
    """
    Detecta variables PRE y POST cuando existen.
    Devuelve estructura interna:
        {"RPE": {"pre": col_pre, "post": col_post}, ...}
    """
    patrones_base = {
        "RPE": ["rpe", "percepcion", "percepciÃ³n", "esfuerzo"],
        "Molestias (MS)": ["ms", "molestia", "molestias", "soreness", "dolor"],
        "Capacidad fÃ­sica (PF)": ["pf", "capacidad fisica", "capacidad fÃ­sica", "physical fitness"],
        "Capacidad mental (MF)": ["mf", "capacidad mental", "mental fitness"],
        "Confianza (RC)": ["rc", "confianza", "confidence"],
    }
    pre_tokens = ["pre", "before", "antes", "previo", "pre_match", "prematch"]
    post_tokens = ["post", "after", "despues", "despuÃ©s", "posterior", "post_match", "postmatch"]

    norm_cols = {c: normalizar(c) for c in df.columns}
    variables = {}

    def score_col(nc, patrones, tokens):
        base_ok = any(normalizar(p) in nc for p in patrones)
        token_ok = any(normalizar(t) in nc for t in tokens)
        return base_ok and token_ok

    for etiqueta, patrones in patrones_base.items():
        pre = None
        post = None
        # 1) Coincidencia explÃ­cita PRE/POST
        for c, nc in norm_cols.items():
            if pre is None and score_col(nc, patrones, pre_tokens):
                if pd.to_numeric(df[c], errors="coerce").notna().sum() > 0:
                    pre = c
            if post is None and score_col(nc, patrones, post_tokens):
                if pd.to_numeric(df[c], errors="coerce").notna().sum() > 0:
                    post = c
        # 2) Si no hay PRE/POST, detectar columna genÃ©rica como POST para compatibilidad
        if pre is None and post is None:
            for c, nc in norm_cols.items():
                if any(normalizar(p) in nc for p in patrones):
                    if pd.to_numeric(df[c], errors="coerce").notna().sum() > 0:
                        post = c
                        break
        if pre is not None or post is not None:
            variables[etiqueta] = {"pre": pre, "post": post}

    print("\nVARIABLES PRE/POST DETECTADAS")
    for k, v in variables.items():
        print(f"- {k}: PRE={v.get('pre') if v.get('pre') else 'No'} | POST={v.get('post') if v.get('post') else 'No'}")

    if not variables:
        print("ERROR: no se detectaron variables numÃ©ricas PRE/POST.")
        sys.exit(1)
    return variables


def convertir_variables(df, variables):
    df = df.copy()
    for info in variables.values():
        for c in [info.get("pre"), info.get("post")]:
            if c and c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def preparar_variables_pre_post(df, variables):
    """
    AÃ±ade columnas de diferencia POST-PRE y genera el diccionario de variables para tablas/grÃ¡ficos.
    """
    df = df.copy()
    variables_analisis = {}
    abrevs = {
        "RPE": "RPE",
        "Molestias (MS)": "MS",
        "Capacidad fÃ­sica (PF)": "PF",
        "Capacidad mental (MF)": "MF",
        "Confianza (RC)": "RC",
    }
    for etiqueta, info in variables.items():
        ab = abrevs.get(etiqueta, nombre_seguro(etiqueta))
        pre = info.get("pre")
        post = info.get("post")
        if pre and pre in df.columns:
            variables_analisis[f"{etiqueta} PRE"] = pre
        if post and post in df.columns:
            variables_analisis[f"{etiqueta} POST"] = post
        if pre and post and pre in df.columns and post in df.columns:
            col_delta = f"Delta_{ab}"
            df[col_delta] = pd.to_numeric(df[post], errors="coerce") - pd.to_numeric(df[pre], errors="coerce")
            variables_analisis[f"Î”{ab}"] = col_delta
    return df, variables_analisis

# =============================================================================
# DATOS ÃRBITRO Y FOTO
# =============================================================================

def buscar_fila_arbitro(df_info, arbitro):
    if df_info is None or df_info.empty:
        return None
    col_nombre = detectar_columna(df_info, ["full name", "nombre completo", "nombre", "arbitro", "Ã¡rbitro", "referee"], posicion=0)
    objetivo = normalizar_simple(arbitro)
    for _, row in df_info.iterrows():
        valor = normalizar_simple(row.get(col_nombre, ""))
        if valor == objetivo:
            return row
    for _, row in df_info.iterrows():
        valor = normalizar_simple(row.get(col_nombre, ""))
        if objetivo in valor or valor in objetivo:
            return row
    return None


def fila_a_diccionario(row):
    if row is None:
        return {}
    d = {}
    for k, v in row.items():
        if pd.isna(v):
            continue
        val = limpiar_texto(v)
        if val:
            d[limpiar_texto(k)] = val
    return d


def buscar_valor_datos(datos_arbitro, etiqueta):
    objetivo = normalizar_simple(etiqueta)
    if not datos_arbitro:
        return ""

    equivalencias = {
        "fullname": ["fullname", "full name", "nombre", "nombrecompleto", "arbitro", "referee"],
        "nacionalidad": ["nacionalidad", "nationality", "pais", "country"],
        "age": ["age", "edad"],
        "sex": ["sex", "sexo", "genero", "gender"],
        "heightcm": ["heightcm", "height", "altura", "alturacm"],
        "weightkg": ["weightkg", "weight", "peso", "pesokg"],
        "yearsofrefereeingexperience": ["yearsofrefereeingexperience", "experience", "experiencia", "anosexperiencia", "anosdeexperiencia", "years"],
        "highestlevelofficiated": ["highestlevelofficiated", "highestlevel", "maximonivel", "nivelmaximo"],
        "currentrefereeinglevelcategory": ["currentrefereeinglevelcategory", "currentlevel", "categoriaactual", "nivelactual"],
    }

    claves_busqueda = equivalencias.get(objetivo, [objetivo])

    norm_map = {normalizar_simple(k): v for k, v in datos_arbitro.items()}

    for cb in claves_busqueda:
        ncb = normalizar_simple(cb)
        if ncb in norm_map:
            return norm_map[ncb]

    for k_norm, v in norm_map.items():
        for cb in claves_busqueda:
            ncb = normalizar_simple(cb)
            if ncb in k_norm or k_norm in ncb:
                return v

    return ""


def ruta_foto_desde_datos(datos_dir, datos_arbitro):
    for k, v in datos_arbitro.items():
        if normalizar(k) in ["foto", "fotografia", "imagen", "ruta_foto", "archivo_foto", "photo", "picture"]:
            p = Path(v)
            if not p.is_absolute():
                p = datos_dir / p
            if p.exists() and p.suffix.lower() in EXTENSIONES_IMAGEN:
                return p
    return None


def buscar_foto(datos_dir, arbitro, datos_arbitro):
    p = ruta_foto_desde_datos(datos_dir, datos_arbitro)
    if p:
        return p

    objetivo = normalizar_simple(arbitro)
    imgs = []
    for ext in EXTENSIONES_IMAGEN:
        imgs.extend(datos_dir.rglob(f"*{ext}"))

    for img in imgs:
        if normalizar_simple(img.stem) == objetivo:
            return img
    for img in imgs:
        stem = normalizar_simple(img.stem)
        if objetivo in stem or stem in objetivo:
            return img

    partes = [normalizar_simple(p) for p in re.split(r"\s+|_|-", arbitro) if len(p) >= 3]
    for img in imgs:
        stem = normalizar_simple(img.stem)
        if partes and all(p in stem for p in partes[:2]):
            return img
    return None


# =============================================================================
# ANÃLISIS
# =============================================================================




def calcular_indices_fatiga(df, variables, indices_seleccionados):
    """
    Calcula Ã­ndices PRE, POST y diferencia POST-PRE cuando existen datos suficientes.

    IGFA simple PRE/POST = [RPE + MS + (10-PF) + (10-MF) + (10-RC)] / 5
    IGFA ponderado PRE/POST = 0.30*RPE + 0.25*MS + 0.20*(10-PF) + 0.15*(10-MF) + 0.10*(10-RC)
    Delta = POST - PRE
    """
    df = df.copy()
    requeridas = ["RPE", "Molestias (MS)", "Capacidad fÃ­sica (PF)", "Capacidad mental (MF)", "Confianza (RC)"]
    faltan = [v for v in requeridas if v not in variables]

    if faltan:
        print("AVISO: no se pueden calcular los Ã­ndices de fatiga. Faltan variables:", ", ".join(faltan))
        return df, {}

    def get_var(label, momento):
        c = variables[label].get(momento)
        if c and c in df.columns:
            return pd.to_numeric(df[c], errors="coerce")
        return None

    nuevos = {}

    for momento in ["pre", "post"]:
        rpe = get_var("RPE", momento)
        ms = get_var("Molestias (MS)", momento)
        pf = get_var("Capacidad fÃ­sica (PF)", momento)
        mf = get_var("Capacidad mental (MF)", momento)
        rc = get_var("Confianza (RC)", momento)
        if any(x is None for x in [rpe, ms, pf, mf, rc]):
            continue

        suf = momento.upper()
        if "IGFA simple" in indices_seleccionados:
            col = f"IGFA_simple_{momento}"
            df[col] = ((rpe + ms + (10 - pf) + (10 - mf) + (10 - rc)) / 5).round(2)
            nuevos[f"IGFA simple {suf}"] = col

        if "IGFA ponderado" in indices_seleccionados:
            col = f"IGFA_ponderado_{momento}"
            df[col] = (0.30*rpe + 0.25*ms + 0.20*(10-pf) + 0.15*(10-mf) + 0.10*(10-rc)).round(2)
            nuevos[f"IGFA ponderado {suf}"] = col

    if "IGFA simple PRE" in nuevos and "IGFA simple POST" in nuevos:
        col = "Delta_IGFA_simple"
        df[col] = (df[nuevos["IGFA simple POST"]] - df[nuevos["IGFA simple PRE"]]).round(2)
        nuevos["Î”IGFA simple"] = col

    if "IGFA ponderado PRE" in nuevos and "IGFA ponderado POST" in nuevos:
        col = "Delta_IGFA_ponderado"
        df[col] = (df[nuevos["IGFA ponderado POST"]] - df[nuevos["IGFA ponderado PRE"]]).round(2)
        nuevos["Î”IGFA ponderado"] = col

    return df, nuevos

def clasificar_igfa(valor, idioma="castellano"):
    try:
        v = float(valor)
    except Exception:
        return tr("Sin dato", idioma)
    if v < 2:
        return tr("Fatiga muy baja", idioma)
    if v < 4:
        return tr("Fatiga baja", idioma)
    if v < 6:
        return tr("Fatiga moderada", idioma)
    if v < 8:
        return tr("Fatiga alta", idioma)
    return tr("Fatiga muy alta", idioma)


def tabla_partidos(df, col_partido, variables, indices_variables=None, idioma="castellano"):
    # Incluye GÃ©nero justo despuÃ©s de Partido si existe en los datos filtrados.
    columnas_base = [col_partido]
    if "_GENERO_STD" in df.columns:
        columnas_base.append("_GENERO_STD")

    indices_variables = indices_variables or {}
    columnas = columnas_base + list(variables.values()) + list(indices_variables.values())
    columnas = [c for c in columnas if c in df.columns]

    out = df[columnas].copy()

    rename = {col_partido: "Partido"}
    if "_GENERO_STD" in out.columns:
        rename["_GENERO_STD"] = "GÃ©nero"
    rename.update({v: k for k, v in variables.items()})
    rename.update({v: k for k, v in indices_variables.items()})
    out = out.rename(columns=rename)

    for c in out.columns:
        if c not in ["Partido", "GÃ©nero"]:
            if c.startswith("IGFA"):
                out[c] = pd.to_numeric(out[c], errors="coerce").round(2)
            else:
                out[c] = pd.to_numeric(out[c], errors="coerce").round(0).astype("Int64")

    if "IGFA simple" in out.columns:
        out["ClasificaciÃ³n IGFA simple"] = out["IGFA simple"].map(lambda x: clasificar_igfa(x, idioma))
    if "IGFA ponderado" in out.columns:
        out["ClasificaciÃ³n IGFA ponderado"] = out["IGFA ponderado"].map(lambda x: clasificar_igfa(x, idioma))

    return traducir_df(out, idioma)

def resumen_por_grupo(df, grupo, variables, nombre_grupo):
    filas = []
    for g, sub in df.groupby(grupo, dropna=False):
        fila = {nombre_grupo: limpiar_texto(g) or "Sin dato", "n": len(sub)}
        for etiqueta, col in variables.items():
            s = pd.to_numeric(sub[col], errors="coerce").dropna()
            fila[f"{etiqueta}_media"] = round(float(s.mean()), 2) if len(s) else np.nan
            fila[f"{etiqueta}_sd"] = round(float(s.std(ddof=1)), 2) if len(s) > 1 else "No aplicable"
        filas.append(fila)
    out = pd.DataFrame(filas)
    orden = {normalizar(v): i for i, v in enumerate(VALORES_COMPETITIVOS)}
    if nombre_grupo in out.columns and out[nombre_grupo].map(lambda x: normalizar(x) in orden).all():
        out["_orden"] = out[nombre_grupo].map(lambda x: orden.get(normalizar(x), 999))
        out = out.sort_values("_orden").drop(columns="_orden")
    return out


def grafico_variable_por_partidos(df, col_partido, etiqueta, col_variable, ruta, idioma="castellano"):
    datos = df[[col_partido, col_variable]].copy()
    datos[col_variable] = pd.to_numeric(datos[col_variable], errors="coerce")
    datos = datos.dropna(subset=[col_variable])
    if datos.empty:
        return None

    x = np.arange(len(datos))
    y = datos[col_variable].values
    xlabels = [etiqueta_x(v, 16) for v in datos[col_partido].astype(str)]

    plt.figure(figsize=(max(8.8, len(datos)*0.72), 4.9))

    # LÃ­nea real: se representan los valores observados sin suavizado ni interpolaciÃ³n.
    plt.plot(x, y, color="black", linewidth=2.2)

    # Puntos: borde negro, interior blanco.
    plt.scatter(x, y, s=52, facecolors="white", edgecolors="black", linewidths=1.7, zorder=3)

    # LÃ­nea de tendencia discontinua con ecuaciÃ³n, sin recuadro.
    tendencia = None
    if len(y) >= 2:
        m, b = np.polyfit(x, y, 1)
        yp = m * x + b
        ss_res = np.sum((y - yp) ** 2)
        ss_tot = np.sum((y - np.mean(y)) ** 2)
        r2 = 1 - ss_res / ss_tot if ss_tot else 0.0
        tendencia = (float(m), float(r2))
        plt.plot(x, yp, color="black", linestyle="--", linewidth=1.25, alpha=0.75)
        signo = "+" if b >= 0 else "-"
        ecuacion = f"y = {m:.2f}x {signo} {abs(b):.2f} | RÂ² = {r2:.2f}"
        plt.gcf().text(0.98, 0.965, ecuacion, ha="right", va="top", fontsize=9, color="black")

    # NÃºmeros sin decimales.
    for xi, yi in zip(x, y):
        plt.text(xi, yi + 0.18, f"{yi:.0f}", ha="center", fontsize=9, color="black")

    plt.xticks(x, xlabels, rotation=0, ha="center", fontsize=8)
    plt.ylim(0, 10.5)
    plt.ylabel(TEXTOS[idioma]["puntuacion_directa"])
    plt.title(f"{etiqueta_idioma(etiqueta, idioma)} {TEXTOS[idioma]['por_partido']}")
    plt.grid(axis="y", alpha=0.22)
    plt.gca().spines["top"].set_visible(False)
    plt.gca().spines["right"].set_visible(False)

    plt.tight_layout(rect=[0, 0, 1, 0.93])
    plt.savefig(ruta, dpi=240)
    plt.close()

    return {
        "min": float(np.nanmin(y)),
        "max": float(np.nanmax(y)),
        "media": float(np.nanmean(y)),
        "n": int(len(y)),
        "tendencia": tendencia,
    }


def grafico_variable_por_grupo(resumen, columna_grupo, etiqueta, ruta, titulo, idioma="castellano"):
    media_col = f"{etiqueta}_media"
    if media_col not in resumen.columns:
        return None

    datos = resumen[[columna_grupo, media_col]].dropna()
    if len(datos) < 2:
        return None

    x = np.arange(len(datos))
    y = datos[media_col].astype(float).values
    xlabels = [etiqueta_x(v, 18) for v in datos[columna_grupo].astype(str)]

    plt.figure(figsize=(8.6, 4.7))

    # LÃ­nea real: se representan los promedios observados por grupo sin suavizado ni interpolaciÃ³n.
    plt.plot(x, y, color="black", linewidth=2.2)

    plt.scatter(x, y, s=52, facecolors="white", edgecolors="black", linewidths=1.7, zorder=3)

    tendencia = None
    if len(y) >= 2:
        m, b = np.polyfit(x, y, 1)
        yp = m*x + b
        ss_res = np.sum((y-yp)**2)
        ss_tot = np.sum((y-np.mean(y))**2)
        r2 = 1 - ss_res/ss_tot if ss_tot else 0.0
        tendencia = (float(m), float(r2))
        plt.plot(x, yp, color="black", linestyle="--", linewidth=1.25, alpha=0.75)
        signo = "+" if b >= 0 else "-"
        plt.gcf().text(0.98, 0.965, f"y = {m:.2f}x {signo} {abs(b):.2f} | RÂ² = {r2:.2f}", ha="right", va="top", fontsize=9, color="black")

    for xi, yi in zip(x, y):
        plt.text(xi, yi + 0.18, f"{yi:.0f}", ha="center", fontsize=9, color="black")

    plt.xticks(x, xlabels, rotation=0, ha="center", fontsize=9)
    plt.ylim(0, 10.5)
    plt.ylabel(TEXTOS[idioma]["media_010"])
    plt.title(titulo)
    plt.grid(axis="y", alpha=0.22)
    plt.gca().spines["top"].set_visible(False)
    plt.gca().spines["right"].set_visible(False)
    plt.tight_layout(rect=[0, 0, 1, 0.93])
    plt.savefig(ruta, dpi=240)
    plt.close()

    return {"media": float(np.nanmean(y)), "min": float(np.nanmin(y)), "max": float(np.nanmax(y)), "n": int(len(y)), "tendencia": tendencia}



def sentido_variable(etiqueta):
    e = normalizar(etiqueta)
    peor_alto = ["rpe", "molestias", "ms", "soreness", "igfa", "grfi"]
    mejor_alto = ["capacidad_fisica", "pf", "capacidad_mental", "mf", "confianza", "rc"]
    if any(k in e for k in peor_alto):
        return "peor_alto"
    if any(k in e for k in mejor_alto):
        return "mejor_alto"
    if e.startswith("delta") or e.startswith("d") or "delta" in e:
        if any(k in e for k in ["rpe", "ms", "igfa", "grfi"]):
            return "delta_peor_alto"
        if any(k in e for k in ["pf", "mf", "rc"]):
            return "delta_mejor_alto"
    return "neutro"


def lectura_tendencia(etiqueta, tendencia, idioma="castellano"):
    if not tendencia:
        return "no pudo estimarse una tendencia lineal estable" if idioma != "ingles" else "a stable linear trend could not be estimated"
    m, r2 = tendencia
    if r2 < 0.10:
        return "la tendencia lineal se muestra con baja capacidad explicativa (RÂ² < 0,10); debe interpretarse con prudencia" if idioma != "ingles" else "the linear trend is shown with low explanatory capacity (RÂ² < 0.10) and should be interpreted cautiously"
    estado = sentido_variable(etiqueta)
    if abs(m) < 0.05:
        return f"la tendencia fue estable (pendiente={m:.2f}; RÂ²={r2:.2f})" if idioma != "ingles" else f"the trend was stable (slope={m:.2f}; RÂ²={r2:.2f})"
    if idioma == "ingles":
        direccion = "upward" if m > 0 else "downward"
        if estado in ["peor_alto", "delta_peor_alto"]:
            significado = "unfavourable" if m > 0 else "favourable"
        elif estado in ["mejor_alto", "delta_mejor_alto"]:
            significado = "favourable" if m > 0 else "unfavourable"
        else:
            significado = "descriptive"
        return f"the trend was {direccion} and {significado} (slope={m:.2f}; RÂ²={r2:.2f})"
    direccion = "ascendente" if m > 0 else "descendente"
    if estado in ["peor_alto", "delta_peor_alto"]:
        significado = "desfavorable" if m > 0 else "favorable"
    elif estado in ["mejor_alto", "delta_mejor_alto"]:
        significado = "favorable" if m > 0 else "desfavorable"
    else:
        significado = "descriptiva"
    return f"la tendencia fue {direccion} y de significado {significado} (pendiente={m:.2f}; RÂ²={r2:.2f})"


def interpretar_partido(etiqueta, stats, idioma="castellano"):
    if not stats:
        return "No hay datos suficientes." if idioma != "ingles" else "There are not enough data."
    estado = sentido_variable(etiqueta)
    if idioma == "ingles":
        sentido = "Higher values indicate a worse subjective state." if estado in ["peor_alto", "delta_peor_alto"] else "Higher values indicate a more favourable subjective state." if estado in ["mejor_alto", "delta_mejor_alto"] else "The variable is interpreted descriptively."
        return (f"{stats['n']} matches were analysed using direct match values, without smoothing or interpolation. "
                f"The mean was {stats['media']:.2f}, with values ranging from {stats['min']:.0f} to {stats['max']:.0f}. "
                f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {sentido}")
    sentido = "Valores altos indican peor estado subjetivo." if estado in ["peor_alto", "delta_peor_alto"] else "Valores altos indican un estado subjetivo mÃ¡s favorable." if estado in ["mejor_alto", "delta_mejor_alto"] else "La variable se interpreta de forma descriptiva."
    return (f"Se analizaron {stats['n']} partidos mediante valores directos, sin suavizado ni interpolaciÃ³n. "
            f"La media fue {stats['media']:.2f}, con valores entre {stats['min']:.0f} y {stats['max']:.0f}. "
            f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {sentido}")


def interpretar_grupo(etiqueta, stats, tipo, idioma="castellano"):
    if not stats:
        return "No hay datos suficientes." if idioma != "ingles" else "There are not enough data."
    estado = sentido_variable(etiqueta)
    if idioma == "ingles":
        sentido = "Higher group means indicate worse subjective response." if estado in ["peor_alto", "delta_peor_alto"] else "Higher group means indicate a more favourable subjective response." if estado in ["mejor_alto", "delta_mejor_alto"] else "Group means are interpreted descriptively."
        return (f"The {tipo} analysis compared observed group means without smoothing. "
                f"The overall mean was {stats['media']:.2f}, with group means ranging from {stats['min']:.0f} to {stats['max']:.0f}. "
                f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {sentido}")
    sentido = "Medias mÃ¡s altas indican peor respuesta subjetiva." if estado in ["peor_alto", "delta_peor_alto"] else "Medias mÃ¡s altas indican una respuesta subjetiva mÃ¡s favorable." if estado in ["mejor_alto", "delta_mejor_alto"] else "Las medias se interpretan de forma descriptiva."
    return (f"El anÃ¡lisis {tipo} comparÃ³ medias observadas por grupo, sin suavizado. "
            f"La media global fue {stats['media']:.2f}, con medias entre {stats['min']:.0f} y {stats['max']:.0f}. "
            f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {sentido}")


def valor_extremo(df, col_partido, col, peor_alto=True):
    if col not in df.columns:
        return None
    s = pd.to_numeric(df[col], errors="coerce")
    if s.dropna().empty:
        return None
    idx = s.idxmax() if peor_alto else s.idxmin()
    return limpiar_texto(df.loc[idx, col_partido]), float(s.loc[idx])


def tabla_resumen_global(df, col_partido, variables_analisis, col_competitiva=None, idioma="castellano"):
    filas = []
    for etiqueta, col in variables_analisis.items():
        s = pd.to_numeric(df[col], errors="coerce").dropna() if col in df.columns else pd.Series(dtype=float)
        if s.empty:
            continue
        estado = sentido_variable(etiqueta)
        peor_alto = estado in ["peor_alto", "delta_peor_alto", "neutro"]
        partido, valor = valor_extremo(df, col_partido, col, peor_alto=peor_alto)
        m = np.nan; r2 = np.nan
        if len(s) >= 2:
            y = pd.to_numeric(df[col], errors="coerce").dropna().values
            x = np.arange(len(y))
            m, b = np.polyfit(x, y, 1)
            yp = m*x + b
            ss_res = np.sum((y-yp)**2); ss_tot = np.sum((y-np.mean(y))**2)
            r2 = 1 - ss_res/ss_tot if ss_tot else 0.0
        filas.append({
            "Variable": etiqueta, "n": int(s.count()), "Media": round(float(s.mean()),2),
            "MÃ­nimo": round(float(s.min()),2), "MÃ¡ximo": round(float(s.max()),2),
            "Partido crÃ­tico": partido, "Valor crÃ­tico": round(float(valor),2),
            "Pendiente": round(float(m),3) if pd.notna(m) else "No aplicable",
            "RÂ²": round(float(r2),3) if pd.notna(r2) else "No aplicable",
            "Tendencia visible": "SÃ­" if pd.notna(r2) else "No",
            "Criterio": "Alto = peor" if estado in ["peor_alto", "delta_peor_alto"] else "Alto = mejor" if estado in ["mejor_alto", "delta_mejor_alto"] else "Descriptivo"
        })
    out = pd.DataFrame(filas)
    return traducir_df(out, idioma)


def hallazgos_principales(df, col_partido, variables_base, variables_analisis, col_competitiva=None, idioma="castellano"):
    h = []
    def add(txt):
        if txt: h.append(txt)
    # partido mÃ¡s exigente por IGFA ponderado/simple POST o RPE POST
    candidatos = ["IGFA ponderado POST", "IGFA simple POST", "RPE POST", "Molestias (MS) POST"]
    for et in candidatos:
        if et in variables_analisis:
            ex = valor_extremo(df, col_partido, variables_analisis[et], peor_alto=True)
            if ex:
                add(f"Partido mÃ¡s exigente: {ex[0]} ({et}={ex[1]:.2f}).")
                break
    for et in ["Î”IGFA ponderado", "Î”IGFA simple", "Î”RPE", "Î”MS"]:
        if et in variables_analisis:
            ex = valor_extremo(df, col_partido, variables_analisis[et], peor_alto=True)
            if ex:
                add(f"Mayor incremento de fatiga: {ex[0]} ({et}={ex[1]:.2f}; valores positivos indican empeoramiento cuando la variable es RPE, MS o IGFA).")
                break
    for et in ["Capacidad fÃ­sica (PF) POST", "Capacidad mental (MF) POST", "Confianza (RC) POST"]:
        if et in variables_analisis:
            ex = valor_extremo(df, col_partido, variables_analisis[et], peor_alto=False)
            if ex:
                add(f"Peor recuperaciÃ³n/disposiciÃ³n percibida: {ex[0]} ({et}={ex[1]:.2f}; valores bajos son menos favorables).")
                break
    for et in ["Confianza (RC) POST", "Confianza (RC) PRE"]:
        if et in variables_analisis:
            s = pd.to_numeric(df[variables_analisis[et]], errors="coerce").dropna()
            if len(s) >= 2:
                cv = float(s.std(ddof=1) / s.mean()) if s.mean() else np.nan
                estado = "alta" if pd.notna(cv) and cv < 0.15 else "moderada" if pd.notna(cv) and cv < 0.30 else "baja"
                add(f"Estabilidad de confianza: {estado} (CV={cv:.2f}) en {et}.")
                break
    # dificultad competitiva
    if col_competitiva and col_competitiva in df.columns:
        counts = df[col_competitiva].astype(str).value_counts()
        if not counts.empty:
            add("AnÃ¡lisis por dificultad competitiva incorporado: " + "; ".join([f"{k}: {v}" for k,v in counts.items()]) + ".")
    # acumulaciÃ³n: pendiente de Ã­ndice/RPE si R2 visible
    for et in ["IGFA ponderado POST", "IGFA simple POST", "RPE POST", "Molestias (MS) POST"]:
        if et in variables_analisis:
            s = pd.to_numeric(df[variables_analisis[et]], errors="coerce").dropna()
            if len(s) >= 3:
                x=np.arange(len(s)); y=s.values; m,b=np.polyfit(x,y,1); yp=m*x+b
                ss_res=np.sum((y-yp)**2); ss_tot=np.sum((y-np.mean(y))**2); r2=1-ss_res/ss_tot if ss_tot else 0.0
                if m > 0 and r2 >= 0.10:
                    add(f"AcumulaciÃ³n de fatiga: compatible con aumento progresivo en {et} (pendiente={m:.2f}; RÂ²={r2:.2f}).")
                elif r2 >= 0.10:
                    add(f"No se observa acumulaciÃ³n creciente de fatiga en {et} (pendiente={m:.2f}; RÂ²={r2:.2f}).")
                else:
                    add(f"No se informa tendencia acumulativa en {et} por RÂ² < 0,10.")
                break
    return h


def interpretacion_integrada(hallazgos, idioma="castellano"):
    if idioma == "ingles":
        return "Integrated interpretation: " + " ".join(hallazgos) if hallazgos else "Integrated interpretation: the available data do not support a robust automatic interpretation."
    return "InterpretaciÃ³n integrada: " + " ".join(hallazgos) if hallazgos else "InterpretaciÃ³n integrada: los datos disponibles no permiten una interpretaciÃ³n automÃ¡tica robusta."


def conclusion_cientifica_automatica(hallazgos, idioma="castellano"):
    if idioma == "ingles":
        return ("Scientific conclusion: the report identifies the most demanding matches and the main subjective fatigue/recovery patterns. "
                "All conclusions are descriptive and depend on the available observations; no inferential claims are made automatically. " + " ".join(hallazgos[:3]))
    return ("ConclusiÃ³n cientÃ­fica: el informe identifica los partidos de mayor exigencia y los principales patrones subjetivos de fatiga, recuperaciÃ³n y confianza. "
            "Las conclusiones son descriptivas y dependen de las observaciones disponibles; no se formulan inferencias automÃ¡ticas no justificadas. " + " ".join(hallazgos[:3]))

# =============================================================================
# WORD
# =============================================================================

def set_cell(cell, text, bold=False, size=8):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("-" if pd.isna(text) else str(text))
    r.bold = bold
    r.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_heading(doc, text, size=14):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = COLOR_TITULO



def add_table_df(doc, df, idioma="castellano"):
    if df.empty:
        doc.add_paragraph(TEXTOS[idioma]["sin_datos"])
        return
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, c in enumerate(df.columns):
        set_cell(table.rows[0].cells[j], c, bold=True, size=8)
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for j, c in enumerate(df.columns):
            val = row[c]
            if isinstance(val, float) and not pd.isna(val):
                val = f"{val:.2f}"
            set_cell(cells[j], val, size=8)
    doc.add_paragraph("")


def tabla_resumen_vertical(df, columna_grupo):
    """
    Cambia orientaciÃ³n de tablas resumen anchas:
    MÃ©trica | grupo_1 | grupo_2 | grupo_3 ...
    """
    if df is None or df.empty or columna_grupo not in df.columns:
        return df
    grupos = [str(x) for x in df[columna_grupo].tolist()]
    metricas = [c for c in df.columns if c != columna_grupo]
    filas = []
    for m in metricas:
        fila = {"MÃ©trica": m}
        for _, row in df.iterrows():
            fila[str(row[columna_grupo])] = row[m]
        filas.append(fila)
    return pd.DataFrame(filas)


def add_table_resumen_vertical(doc, df, columna_grupo, idioma="castellano"):
    add_table_df(doc, tabla_resumen_vertical(df, columna_grupo), idioma=idioma)


def tabla_siglas_variables(idioma="castellano"):
    if idioma == "ingles":
        return pd.DataFrame([
            {"Abbreviation": "RPE", "Full name": "Rating of Perceived Exertion", "Brief explanation": "Subjective indicator of perceived exercise intensity. It is useful for monitoring internal load because it integrates physiological and perceptual signals from the athlete or referee (Foster et al., 2001; Saw et al., 2016)."},
            {"Abbreviation": "MS", "Full name": "Muscle Soreness", "Brief explanation": "Reflects perceived muscle soreness or discomfort. It helps contextualise recovery status and potential accumulated fatigue responses after repeated efforts (Saw et al., 2016; Kellmann et al., 2018)."},
            {"Abbreviation": "PF", "Full name": "Physical Fitness", "Brief explanation": "Describes the perceived physical state available for competition. It can complement external load and help contextualise functional readiness before or after the match (Saw et al., 2016; McLaren et al., 2018)."},
            {"Abbreviation": "MF", "Full name": "Mental Fitness", "Brief explanation": "Summarises perceived mental clarity, concentration and readiness. It is relevant because mental fatigue can increase perceived exertion and affect decision-making performance (Marcora et al., 2009)."},
            {"Abbreviation": "RC", "Full name": "Referee Confidence", "Brief explanation": "Represents perceived confidence to officiate and make decisions during the match. In competitive contexts, confidence may modulate psychological response and officiating performance (GuillÃ©n & Feltz, 2011)."},
            {"Abbreviation": "Simple GRFI", "Full name": "Simple Global Referee Fatigue Index", "Brief explanation": "Composite index calculated as the mean of RPE, muscle soreness and the inverted positive variables: 10-PF, 10-MF and 10-RC. It summarises global subjective fatigue on a 0-10 scale. Its rationale is based on the combined use of subjective measures to monitor internal response and recovery (Saw et al., 2016; Kellmann et al., 2018)."},
            {"Abbreviation": "Weighted GRFI", "Full name": "Weighted Global Referee Fatigue Index", "Brief explanation": "Composite index giving greater weight to RPE and muscle soreness: 0.30Â·RPE + 0.25Â·MS + 0.20Â·(10-PF) + 0.15Â·(10-MF) + 0.10Â·(10-RC). It prioritises perceived load and muscle discomfort while maintaining physical, mental and confidence components as subjective modulators (Foster et al., 2001; McLaren et al., 2018)."},
        ])

    return pd.DataFrame([
        {"Sigla": "RPE", "Nombre completo": "Rating of Perceived Exertion / PercepciÃ³n subjetiva del esfuerzo", "ExplicaciÃ³n breve": "Indicador subjetivo de la intensidad percibida. Es Ãºtil para monitorizar la carga interna porque integra seÃ±ales fisiolÃ³gicas y perceptivas del deportista o Ã¡rbitro (Foster et al., 2001; Saw et al., 2016)."},
        {"Sigla": "MS", "Nombre completo": "Muscle Soreness / Molestias musculares", "ExplicaciÃ³n breve": "Refleja la percepciÃ³n de molestias o dolor muscular. Ayuda a interpretar el estado de recuperaciÃ³n y posibles respuestas de fatiga acumulada tras esfuerzos repetidos (Saw et al., 2016; Kellmann et al., 2018)."},
        {"Sigla": "PF", "Nombre completo": "Physical Fitness / Capacidad fÃ­sica percibida", "ExplicaciÃ³n breve": "Describe la percepciÃ³n del estado fÃ­sico disponible para competir. Puede complementar la carga externa y ayudar a contextualizar la preparaciÃ³n funcional antes o despuÃ©s del partido (Saw et al., 2016; McLaren et al., 2018)."},
        {"Sigla": "MF", "Nombre completo": "Mental Fitness / Capacidad mental percibida", "ExplicaciÃ³n breve": "Resume la percepciÃ³n de claridad, concentraciÃ³n y disposiciÃ³n mental. Es relevante porque la fatiga mental puede aumentar el esfuerzo percibido y afectar al rendimiento decisional (Marcora et al., 2009)."},
        {"Sigla": "RC", "Nombre completo": "Referee Confidence / Confianza arbitral percibida", "ExplicaciÃ³n breve": "Representa la seguridad percibida para actuar y tomar decisiones durante el partido. En contextos competitivos, la confianza puede modular la respuesta psicolÃ³gica y la calidad de la actuaciÃ³n (GuillÃ©n & Feltz, 2011)."},
        {"Sigla": "IGFA simple", "Nombre completo": "Ãndice Global de Fatiga Arbitral simple", "ExplicaciÃ³n breve": "Ãndice compuesto calculado como la media de RPE, molestias y las variables positivas invertidas: 10-PF, 10-MF y 10-RC. Resume la fatiga subjetiva global en escala 0-10. Su lÃ³gica se apoya en el uso combinado de medidas subjetivas para monitorizar la respuesta interna y la recuperaciÃ³n (Saw et al., 2016; Kellmann et al., 2018)."},
        {"Sigla": "IGFA ponderado", "Nombre completo": "Ãndice Global de Fatiga Arbitral ponderado", "ExplicaciÃ³n breve": "Ãndice compuesto que otorga mayor peso a RPE y molestias: 0.30Â·RPE + 0.25Â·MS + 0.20Â·(10-PF) + 0.15Â·(10-MF) + 0.10Â·(10-RC). Permite priorizar la carga percibida y el malestar muscular, manteniendo el componente fÃ­sico, mental y de confianza como moduladores subjetivos (Foster et al., 2001; McLaren et al., 2018)."},

        {"Sigla": "Dificultad competitiva", "Nombre completo": "Nivel competitivo del partido", "ExplicaciÃ³n breve": "Indica si un partido fue muy igualado o muy desequilibrado. Se calcula principalmente a partir de la diferencia final del marcador, incorporando ademÃ¡s el volumen total de puntos y la posiciÃ³n relativa del partido dentro del torneo. Una dificultad alta representa encuentros mÃ¡s igualados y competitivos; una dificultad baja representa partidos con diferencias amplias en el marcador."},
    ])


def insertar_tabla_siglas_y_referencias(doc, idioma="castellano"):
    add_heading(doc, TEXTOS[idioma]["variables"], size=14)
    doc.add_paragraph(TEXTOS[idioma]["variables_intro"])
    add_table_df(doc, tabla_siglas_variables(idioma), idioma=idioma)

    add_heading(doc, TEXTOS[idioma]["referencias"], size=14)
    refs = [
        "Foster, C., Florhaug, J. A., Franklin, J., Gottschall, L., Hrovatin, L. A., Parker, S., Doleshal, P., & Dodge, C. (2001). A new approach to monitoring exercise training. Journal of Strength and Conditioning Research, 15(1), 109â€“115.",
        "GuillÃ©n, F., & Feltz, D. L. (2011). A conceptual model of referee efficacy. Frontiers in Psychology, 2, 25.",
        "Kellmann, M., Bertollo, M., Bosquet, L., Brink, M., Coutts, A. J., Duffield, R., Erlacher, D., Halson, S. L., Hecksteden, A., Heidari, J., Kallus, K. W., Meeusen, R., Mujika, I., Robazza, C., Skorski, S., Venter, R., & Beckmann, J. (2018). Recovery and performance in sport: Consensus statement. International Journal of Sports Physiology and Performance, 13(2), 240â€“245.",
        "Marcora, S. M., Staiano, W., & Manning, V. (2009). Mental fatigue impairs physical performance in humans. Journal of Applied Physiology, 106(3), 857â€“864.",
        "McLaren, S. J., Macpherson, T. W., Coutts, A. J., Hurst, C., Spears, I. R., & Weston, M. (2018). The relationships between internal and external measures of training load and intensity in team sports: A meta-analysis. Sports Medicine, 48(3), 641â€“658.",
        "Saw, A. E., Main, L. C., & Gastin, P. B. (2016). Monitoring the athlete training response: Subjective self-reported measures trump commonly used objective measures. British Journal of Sports Medicine, 50(5), 281â€“291.",
    ]
    for ref in refs:
        p = doc.add_paragraph(ref)
        for run in p.runs:
            run.font.size = Pt(8)

def rellenar_tabla_datos_y_foto(doc, datos_arbitro, ruta_foto):
    """
    Busca la tabla de la plantilla que contiene Variable / Valor / Foto.
    Rellena la columna Valor y sustituye la celda Foto por la imagen.
    """
    tabla_obj = None
    for table in doc.tables:
        textos = " ".join(cell.text for row in table.rows for cell in row.cells)
        nt = normalizar_simple(textos)
        if "variable" in nt and "valor" in nt and "foto" in nt:
            tabla_obj = table
            break

    if tabla_obj is None:
        print("AVISO: no se encontrÃ³ la tabla Variable/Valor/Foto en la plantilla.")
        return False

    foto_insertada = False

    # Rellenar columna Valor: normalmente columna 0 = Variable, columna 1 = Valor.
    for row in tabla_obj.rows:
        if len(row.cells) < 2:
            continue
        etiqueta = limpiar_texto(row.cells[0].text)
        if normalizar_simple(etiqueta) in ["variable", ""]:
            continue

        valor = buscar_valor_datos(datos_arbitro, etiqueta)
        if valor:
            set_cell(row.cells[1], valor, bold=False, size=9)

        # Celda foto: se detecta cualquier celda con texto Foto.
        for cell in row.cells:
            if normalizar_simple(cell.text) == "foto" and not foto_insertada:
                cell.text = ""
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if ruta_foto and Path(ruta_foto).exists():
                    run = p.add_run()
                    run.add_picture(str(ruta_foto), width=Inches(2.0))
                    foto_insertada = True
                else:
                    p.add_run("Sin fotografÃ­a detectada.")

    # Por si la celda foto estÃ¡ combinada y solo aparece una vez fuera del loop Ãºtil.
    if not foto_insertada:
        for row in tabla_obj.rows:
            for cell in row.cells:
                if "foto" in normalizar_simple(cell.text):
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if ruta_foto and Path(ruta_foto).exists():
                        run = p.add_run()
                        run.add_picture(str(ruta_foto), width=Inches(2.0))
                        foto_insertada = True
                    else:
                        p.add_run("Sin fotografÃ­a detectada.")
                    break
            if foto_insertada:
                break

    return True





# =============================================================================
# V30 - FUNCIONES PROFESIONALES REESCRITAS
# =============================================================================

def limpiar_num_epigrafe(txt):
    txt = limpiar_texto(txt)
    return re.sub(r"^\s*\d+\s*[\.-]\s*", "", txt).strip()


def sentido_variable(etiqueta):
    """
    Criterio conceptual unificado.
    Variables de carga/fatiga: valores altos = peor estado.
    Variables positivas: valores altos = mejor estado.
    Deltas: siempre POST - PRE.
      - Î”RPE, Î”MS, Î”IGFA > 0 = empeoramiento.
      - Î”PF, Î”MF, Î”RC > 0 = mejora/disposiciÃ³n mÃ¡s favorable.
    """
    e_raw = str(etiqueta).strip().lower()
    e = normalizar(etiqueta)
    es_delta = e_raw.startswith("Î´") or e_raw.startswith("Î”".lower()) or e.startswith("delta") or "delta" in e
    peor = any(k in e for k in ["rpe", "molestias", "ms", "soreness", "igfa", "grfi"])
    mejor = any(k in e for k in ["capacidad_fisica", "physical_fitness", "pf", "capacidad_mental", "mental_fitness", "mf", "confianza", "confidence", "rc"])
    if es_delta and peor:
        return "delta_peor_alto"
    if es_delta and mejor:
        return "delta_mejor_alto"
    if peor:
        return "peor_alto"
    if mejor:
        return "mejor_alto"
    return "neutro"


def criterio_delta_texto(etiqueta, idioma="castellano"):
    estado = sentido_variable(etiqueta)
    if estado == "delta_peor_alto":
        return ("Diferencia POST-PRE: valores positivos indican aumento de carga/fatiga percibida; valores negativos indican reducciÃ³n o mejor recuperaciÃ³n." if idioma != "ingles" else
                "POST-PRE difference: positive values indicate increased perceived load/fatigue; negative values indicate reduction or better recovery.")
    if estado == "delta_mejor_alto":
        return ("Diferencia POST-PRE: valores positivos indican mejora de la disposiciÃ³n percibida; valores negativos indican deterioro del estado fÃ­sico, mental o de confianza." if idioma != "ingles" else
                "POST-PRE difference: positive values indicate improved perceived readiness; negative values indicate deterioration in physical, mental or confidence status.")
    return ""


def buscar_foto(datos_dir, arbitro, datos_arbitro):
    """BÃºsqueda robusta de fotografÃ­a del Ã¡rbitro en Datos y subcarpetas."""
    p = ruta_foto_desde_datos(datos_dir, datos_arbitro)
    if p and p.exists():
        return p

    objetivo = normalizar_simple(arbitro)
    partes = [normalizar_simple(x) for x in re.split(r"\s+|_|-|\.", str(arbitro)) if len(normalizar_simple(x)) >= 3]
    imgs = []
    for ext in EXTENSIONES_IMAGEN:
        imgs.extend(datos_dir.rglob(f"*{ext}"))
        imgs.extend(datos_dir.rglob(f"*{ext.upper()}"))
    imgs = [x for x in imgs if x.exists()]
    if not imgs:
        return None

    # Coincidencias exactas y parciales.
    for img in imgs:
        stem = normalizar_simple(img.stem)
        if stem == objetivo:
            return img
    for img in imgs:
        stem = normalizar_simple(img.stem)
        if objetivo and (objetivo in stem or stem in objetivo):
            return img
    for img in imgs:
        stem = normalizar_simple(img.stem)
        if partes and all(p in stem for p in partes[:2]):
            return img
    for img in imgs:
        stem = normalizar_simple(img.stem)
        if partes and any(p in stem for p in partes):
            return img

    # AproximaciÃ³n final por similitud de texto.
    try:
        from difflib import SequenceMatcher
        scored = []
        for img in imgs:
            stem = normalizar_simple(img.stem)
            scored.append((SequenceMatcher(None, objetivo, stem).ratio(), img))
        scored.sort(reverse=True, key=lambda x: x[0])
        if scored and scored[0][0] >= 0.55:
            return scored[0][1]
    except Exception:
        pass
    return None


def insertar_foto_en_documento(doc, ruta_foto, ancho=2.1):
    """Inserta la foto en la primera celda que contenga 'Foto'. Si no existe, crea un bloque de foto."""
    if ruta_foto is None or not Path(ruta_foto).exists():
        return False
    insertada = False
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "foto" in normalizar_simple(cell.text):
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    try:
                        run.add_picture(str(ruta_foto), width=Inches(ancho))
                        insertada = True
                        return True
                    except Exception:
                        return False
    if not insertada:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        try:
            r.add_picture(str(ruta_foto), width=Inches(ancho))
            return True
        except Exception:
            return False
    return insertada


def rellenar_tabla_datos_y_foto(doc, datos_arbitro, ruta_foto):
    tabla_obj = None
    for table in doc.tables:
        textos = " ".join(cell.text for row in table.rows for cell in row.cells)
        nt = normalizar_simple(textos)
        if "variable" in nt and "valor" in nt:
            tabla_obj = table
            break

    if tabla_obj is not None:
        for row in tabla_obj.rows:
            if len(row.cells) < 2:
                continue
            etiqueta = limpiar_texto(row.cells[0].text)
            if normalizar_simple(etiqueta) in ["variable", ""]:
                continue
            valor = buscar_valor_datos(datos_arbitro, etiqueta)
            if valor:
                set_cell(row.cells[1], valor, bold=False, size=9)
    return insertar_foto_en_documento(doc, ruta_foto, ancho=2.1)


def interpretar_partido(etiqueta, stats, idioma="castellano"):
    if not stats:
        return "No hay datos suficientes."
    estado = sentido_variable(etiqueta)
    delta_txt = criterio_delta_texto(etiqueta, idioma)
    if idioma == "ingles":
        criterio = {
            "peor_alto": "Higher values represent a worse subjective state.",
            "mejor_alto": "Higher values represent a more favourable subjective state.",
            "delta_peor_alto": "Positive POST-PRE values represent worsening.",
            "delta_mejor_alto": "Positive POST-PRE values represent improvement."
        }.get(estado, "Descriptive variable.")
        return (f"Direct observed match values were analysed (n={stats['n']}); no smoothing or interpolation was applied. "
                f"The observed range was {stats['min']:.2f}-{stats['max']:.2f}, with a mean of {stats['media']:.2f}. "
                f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {criterio} {delta_txt}").strip()
    criterio = {
        "peor_alto": "Valores altos representan peor estado subjetivo.",
        "mejor_alto": "Valores altos representan un estado subjetivo mÃ¡s favorable.",
        "delta_peor_alto": "Valores positivos POST-PRE representan empeoramiento.",
        "delta_mejor_alto": "Valores positivos POST-PRE representan mejora."
    }.get(estado, "Variable descriptiva.")
    return (f"Se analizaron valores observados directos por partido (n={stats['n']}), sin suavizado ni interpolaciÃ³n. "
            f"El rango observado fue {stats['min']:.2f}-{stats['max']:.2f}, con una media de {stats['media']:.2f}. "
            f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {criterio} {delta_txt}").strip()


def interpretar_grupo(etiqueta, stats, tipo, idioma="castellano"):
    if not stats:
        return "No hay datos suficientes."
    estado = sentido_variable(etiqueta)
    delta_txt = criterio_delta_texto(etiqueta, idioma)
    if idioma == "ingles":
        return (f"The {tipo} comparison summarises real group means, not modelled or smoothed values. "
                f"Group means ranged from {stats['min']:.2f} to {stats['max']:.2f}. "
                f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {delta_txt}").strip()
    criterio = "medias mÃ¡s altas son desfavorables" if estado in ["peor_alto", "delta_peor_alto"] else "medias mÃ¡s altas son favorables" if estado in ["mejor_alto", "delta_mejor_alto"] else "lectura descriptiva"
    return (f"La comparaciÃ³n {tipo} resume medias reales por grupo, no valores modelizados ni suavizados. "
            f"Las medias oscilaron entre {stats['min']:.2f} y {stats['max']:.2f}; {criterio}. "
            f"{lectura_tendencia(etiqueta, stats.get('tendencia'), idioma)}. {delta_txt}").strip()


def tabla_comparativa_dificultad(df, col_competitiva, variables_analisis, idioma="castellano"):
    if not col_competitiva or col_competitiva not in df.columns:
        return pd.DataFrame()
    filas = []
    orden = {normalizar(v): i for i, v in enumerate(VALORES_COMPETITIVOS)}
    for cat, sub in df.groupby(col_competitiva, dropna=False):
        fila = {"Dificultad competitiva": limpiar_texto(cat) or "Sin dato", "n partidos": int(len(sub))}
        for etiqueta, col in variables_analisis.items():
            if col not in sub.columns:
                continue
            s = pd.to_numeric(sub[col], errors="coerce").dropna()
            if len(s):
                fila[f"{etiqueta} media"] = round(float(s.mean()), 2)
                fila[f"{etiqueta} SD"] = round(float(s.std(ddof=1)), 2) if len(s) > 1 else "No aplicable"
        filas.append(fila)
    out = pd.DataFrame(filas)
    if not out.empty:
        out["_orden"] = out["Dificultad competitiva"].map(lambda x: orden.get(normalizar(x), 999))
        out = out.sort_values("_orden").drop(columns="_orden")
    return traducir_df(out, idioma)


def grafico_heatmap_partido_variables(df, col_partido, variables_analisis, ruta, idioma="castellano"):
    cols = [(et, c) for et, c in variables_analisis.items() if c in df.columns]
    if len(cols) < 2 or len(df) < 2:
        return None
    mat = []
    labels_var = []
    for et, c in cols:
        s = pd.to_numeric(df[c], errors="coerce")
        if s.notna().sum() == 0:
            continue
        v = s.astype(float).values
        mn, mx = np.nanmin(v), np.nanmax(v)
        if np.isclose(mx, mn):
            z = np.full_like(v, 0.5, dtype=float)
        else:
            z = (v - mn) / (mx - mn)
        # Para variables positivas, se invierte para que mÃ¡s oscuro/alto represente peor estado relativo.
        if sentido_variable(et) in ["mejor_alto", "delta_mejor_alto"]:
            z = 1 - z
        mat.append(z)
        labels_var.append(etiqueta_idioma(et, idioma))
    if len(mat) < 2:
        return None
    data = np.vstack(mat).T
    fig_w = max(9, len(labels_var)*0.62)
    fig_h = max(4.8, len(df)*0.42)
    plt.figure(figsize=(fig_w, fig_h))
    im = plt.imshow(data, aspect="auto", vmin=0, vmax=1)
    plt.colorbar(im, fraction=0.025, pad=0.02, label="Carga relativa" if idioma != "ingles" else "Relative burden")
    plt.yticks(np.arange(len(df)), [etiqueta_x(x, 22) for x in df[col_partido].astype(str)], fontsize=8)
    plt.xticks(np.arange(len(labels_var)), labels_var, rotation=45, ha="right", fontsize=8)
    plt.title("Heatmap Partido Ã— Variables" if idioma != "ingles" else "Match Ã— Variables heatmap")
    plt.tight_layout()
    plt.savefig(ruta, dpi=240)
    plt.close()
    return {"n_partidos": int(len(df)), "n_variables": int(len(labels_var))}


def discutir_partidos_criticos(df, col_partido, variables_analisis, col_competitiva=None, idioma="castellano"):
    filas = []
    for etiqueta, col in variables_analisis.items():
        if col not in df.columns:
            continue
        estado = sentido_variable(etiqueta)
        peor_alto = estado in ["peor_alto", "delta_peor_alto", "neutro"]
        ex = valor_extremo(df, col_partido, col, peor_alto=peor_alto)
        if not ex:
            continue
        partido, valor = ex
        cat = "No disponible"
        if col_competitiva and col_competitiva in df.columns:
            sub = df[df[col_partido].astype(str) == str(partido)]
            if not sub.empty:
                cat = limpiar_texto(sub.iloc[0].get(col_competitiva, "")) or "No disponible"
        interpret = criterio_delta_texto(etiqueta, idioma) or ("Valor alto desfavorable" if peor_alto else "Valor bajo desfavorable")
        filas.append({"Variable": etiqueta, "Partido crÃ­tico": partido, "Dificultad competitiva": cat, "Valor": round(valor,2), "Lectura": interpret})
    return traducir_df(pd.DataFrame(filas), idioma)


def discusion_integrada_una_pagina(df, col_partido, variables_analisis, col_competitiva=None, hallazgos=None, idioma="castellano"):
    hallazgos = hallazgos or []
    n = len(df)
    cats = []
    if col_competitiva and col_competitiva in df.columns:
        vc = df[col_competitiva].astype(str).value_counts()
        cats = [f"{k} ({v})" for k, v in vc.items()]
    # variable de fatiga principal
    principal = None
    for et in ["IGFA ponderado POST", "IGFA simple POST", "RPE POST", "Molestias (MS) POST"]:
        if et in variables_analisis:
            principal = et; break
    tendencia_txt = ""
    if principal:
        col = variables_analisis[principal]
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(s) >= 3:
            x = np.arange(len(s)); y = s.values
            m,b = np.polyfit(x,y,1); yp=m*x+b
            ss_res=np.sum((y-yp)**2); ss_tot=np.sum((y-np.mean(y))**2)
            r2=1-ss_res/ss_tot if ss_tot else 0.0
            if r2 >= 0.10:
                tendencia_txt = f"En la variable principal ({principal}) se observÃ³ una pendiente de {m:.2f} con RÂ²={r2:.2f}, por lo que la tendencia se considera interpretable a nivel descriptivo. "
            else:
                tendencia_txt = f"En la variable principal ({principal}) no se informa una tendencia lineal porque RÂ² fue inferior a 0,10. "
    criticos = discutir_partidos_criticos(df, col_partido, variables_analisis, col_competitiva, idioma="castellano")
    crit_txt = ""
    if not criticos.empty:
        top = criticos.head(5)
        crit_txt = "Los partidos crÃ­ticos se concentraron en: " + "; ".join([f"{r['Partido crÃ­tico']} ({r['Variable']}={r['Valor']})" for _, r in top.iterrows()]) + ". "
    dif_txt = ""
    if cats:
        dif_txt = "La distribuciÃ³n por dificultad competitiva fue: " + "; ".join(cats) + ". "
    base = (
        f"La lectura integrada se realizÃ³ sobre {n} observaciones vÃ¡lidas del Ã¡rbitro, combinando variables PRE, POST y diferencias POST-PRE. "
        f"El anÃ¡lisis evita inferencias no justificadas y se centra en patrones descriptivos: magnitud de la respuesta subjetiva, direcciÃ³n del cambio, identificaciÃ³n de partidos crÃ­ticos y relaciÃ³n con la dificultad competitiva. "
        f"{dif_txt}"
        f"{tendencia_txt}"
        f"{crit_txt}"
        f"La interpretaciÃ³n de los deltas es especÃ­fica: en RPE, molestias e IGFA, un delta positivo representa mayor fatiga o peor respuesta postpartido; en capacidad fÃ­sica, capacidad mental y confianza, un delta positivo representa mejora de la disposiciÃ³n percibida, mientras que un delta negativo indica deterioro. "
        f"Desde una perspectiva aplicada, los partidos identificados como crÃ­ticos no deben interpretarse de forma aislada, sino como seÃ±ales para revisar carga acumulada, recuperaciÃ³n entre partidos, contexto competitivo y estabilidad psicolÃ³gica. "
        f"La comparaciÃ³n por dificultad competitiva permite valorar si los encuentros mÃ¡s equilibrados concentran mayor exigencia subjetiva o si la fatiga aparece por acumulaciÃ³n independientemente del equilibrio del marcador. "
        f"Por tanto, el informe debe utilizarse como una herramienta de seguimiento individual del Ã¡rbitro, orientada a detectar momentos de mayor vulnerabilidad, ajustar estrategias de recuperaciÃ³n y apoyar la toma de decisiones durante torneos con congestiÃ³n competitiva."
    )
    if hallazgos:
        base += " Hallazgos automÃ¡ticos principales: " + " ".join(hallazgos[:5])
    if idioma == "ingles":
        return base  # Se mantiene versiÃ³n castellana si se desea una traducciÃ³n manual precisa.
    return base

def crear_informe(ruta_plantilla, ruta_salida, datos_arbitro, ruta_foto, filtros, tabla_p, resumenes, figuras, tipos, n_obs, resumen_global=None, hallazgos=None, interpretacion=None, conclusion_auto=None, tabla_dificultad=None, tabla_criticos=None, discusion_integrada=None, idioma="castellano"):
    doc = Document(str(ruta_plantilla))
    rellenar_tabla_datos_y_foto(doc, datos_arbitro, ruta_foto)

    def heading(titulo, size=14):
        add_heading(doc, limpiar_num_epigrafe(titulo), size=size)

    add_heading(doc, TEXTOS[idioma]["informe"], size=15)
    doc.add_paragraph(f"{TEXTOS[idioma]['observaciones']}: {n_obs}")
    doc.add_paragraph(f"{TEXTOS[idioma]['tipo_analisis_incluido']}: " + ", ".join([tr(t, idioma) for t in sorted(tipos)]))

    heading(TEXTOS[idioma]["filtros"])
    for k, v in filtros.items():
        kk = tr(k, idioma)
        vv = traducir_valor(v, idioma)
        doc.add_paragraph(f"{kk}: {', '.join(vv) if isinstance(vv, list) else vv}")

    if hallazgos:
        heading("Hallazgos principales" if idioma != "ingles" else "Main findings")
        for h in hallazgos:
            try:
                doc.add_paragraph(str(h), style="List Bullet")
            except Exception:
                doc.add_paragraph("â€¢ " + str(h))

    if resumen_global is not None and not resumen_global.empty:
        heading("Tabla resumen cientÃ­fica" if idioma != "ingles" else "Scientific summary table")
        add_table_df(doc, resumen_global, idioma=idioma)

    if "partidos" in tipos:
        heading(TEXTOS[idioma]["partidos"])
        doc.add_paragraph(TEXTOS[idioma]["partidos_intro"])
        add_table_df(doc, tabla_p, idioma=idioma)

    if tabla_dificultad is not None and not tabla_dificultad.empty:
        heading("ComparaciÃ³n real por dificultad competitiva" if idioma != "ingles" else "Real comparison by competitive difficulty")
        doc.add_paragraph("La tabla compara los valores reales observados por categorÃ­a de dificultad competitiva. No se suavizan ni se modelizan los datos; cada celda resume la media y, cuando es posible, la desviaciÃ³n estÃ¡ndar dentro de cada nivel.")
        add_table_resumen_vertical(doc, tabla_dificultad, tr("Dificultad competitiva", idioma), idioma=idioma)

    if tabla_criticos is not None and not tabla_criticos.empty:
        heading("DiscusiÃ³n especÃ­fica de partidos crÃ­ticos" if idioma != "ingles" else "Specific discussion of critical matches")
        doc.add_paragraph("Esta tabla identifica, para cada variable, el partido que representa la situaciÃ³n mÃ¡s comprometida segÃºn el criterio conceptual de interpretaciÃ³n de cada indicador.")
        add_table_df(doc, tabla_criticos, idioma=idioma)

    if figuras:
        heading("Figuras resumen e interpretaciÃ³n analÃ­tica" if idioma != "ingles" else "Summary figures and analytical interpretation")
        for fig in figuras:
            p = doc.add_paragraph()
            r = p.add_run(fig["titulo"])
            r.bold = True
            doc.add_picture(fig["ruta"], width=Inches(6.7))
            inter = fig.get("interpretacion", "")
            if inter:
                doc.add_paragraph(f"{TEXTOS[idioma]['interpretacion']}: " + inter)

    heading("DiscusiÃ³n integrada" if idioma != "ingles" else "Integrated discussion")
    doc.add_paragraph(discusion_integrada or interpretacion or "No se pudo generar una discusiÃ³n integrada por ausencia de datos suficientes.")

    insertar_tabla_siglas_y_referencias(doc, idioma=idioma)
    doc.save(str(ruta_salida))

# =============================================================================
# MAIN
# =============================================================================

def main():
    modo = seleccionar_modo_ejecucion()
    print("="*80)
    print("SCRIPT DE INFORME AUTOMÃTICO PARA ÃRBITRO DE RUGBY 7")
    print("V30 - FOTO ROBUSTA + EPÃGRAFES SECUENCIALES + HEATMAP + DISCUSIÃ“N INTEGRADA")
    print("="*80)

    datos = carpeta_datos(modo)
    ruta_excel = buscar_excel_principal(datos)
    ruta_plantilla = buscar_plantilla(datos)
    ruta_info = buscar_excel_datos_arbitros(datos)

    df = leer_excel(ruta_excel)
    cols = detectar_columnas(df)
    variables_base = detectar_variables(df)
    df = convertir_variables(df, variables_base)
    df, variables = preparar_variables_pre_post(df, variables_base)

    df_info = leer_hoja_opcional(ruta_excel, ["REFEREES", "referees", "arbitros", "Ã¡rbitros", "datos_arbitros"])
    if df_info is None and ruta_info:
        df_info = leer_excel(ruta_info)

    arbitros = valores_unicos(df, cols["arbitro"])
    seleccion_arbitros = seleccionar("ÃRBITROS DETECTADOS EN EL EXCEL", arbitros, permitir_todos=True)
    idiomas = seleccionar_idioma()
    formatos = seleccionar_formato()

    resultados = Path.cwd() / "Resultados"
    resultados.mkdir(exist_ok=True)

    for arbitro in seleccion_arbitros:
        print("\n" + "-"*80)
        print(f"ÃRBITRO: {arbitro}")
        df_a = filtrar(df, cols["arbitro"], [arbitro])

        partidos = valores_unicos(df_a, cols["partido"])
        seleccion_partidos = seleccionar("PARTIDOS DISPONIBLES PARA ESTE ÃRBITRO", partidos, permitir_todos=True)
        df_a = filtrar(df_a, cols["partido"], seleccion_partidos)

        seleccion_comp = []
        if cols.get("competitiva"):
            ops = valores_unicos(df_a, cols["competitiva"])
            if ops:
                seleccion_comp = seleccionar("CATEGORÃA COMPETITIVA", ops, permitir_todos=True)
                df_a = filtrar(df_a, cols["competitiva"], seleccion_comp)

        disponibles = {"partidos": True, "genero": False, "competitiva": cols.get("competitiva") is not None}
        tipos = seleccionar_tipo_analisis(disponibles)
        indices_seleccionados = seleccionar_indices()
        df_a, indices_variables = calcular_indices_fatiga(df_a, variables_base, indices_seleccionados)

        if df_a.empty:
            print("Sin datos tras filtros. Se omite.")
            continue

        variables_analisis = dict(variables)
        variables_analisis.update(indices_variables)

        nombre = nombre_seguro(arbitro)
        out = resultados / f"resultados_{nombre}"
        out.mkdir(parents=True, exist_ok=True)

        fila = buscar_fila_arbitro(df_info, arbitro) if df_info is not None else None
        datos_arbitro = fila_a_diccionario(fila)
        foto = buscar_foto(datos, arbitro, datos_arbitro)
        print("Datos Ã¡rbitro:", "encontrados" if datos_arbitro else "no encontrados")
        print("Foto:", foto.name if foto else "no encontrada")

        resumenes = {}
        if cols.get("competitiva"):
            resumenes["Resumen por categorÃ­a competitiva"] = resumen_por_grupo(df_a, cols["competitiva"], variables_analisis, "CategorÃ­a competitiva")

        resumen_global_es = tabla_resumen_global(df_a, cols["partido"], variables_analisis, cols.get("competitiva"), idioma="castellano")
        tabla_dificultad_es = tabla_comparativa_dificultad(df_a, cols.get("competitiva"), variables_analisis, idioma="castellano")
        tabla_criticos_es = discutir_partidos_criticos(df_a, cols["partido"], variables_analisis, cols.get("competitiva"), idioma="castellano")
        hallazgos_es = hallazgos_principales(df_a, cols["partido"], variables_base, variables_analisis, cols.get("competitiva"), idioma="castellano")
        discusion_es = discusion_integrada_una_pagina(df_a, cols["partido"], variables_analisis, cols.get("competitiva"), hallazgos_es, idioma="castellano")

        tabla_p_es = tabla_partidos(df_a, cols["partido"], variables, indices_variables, idioma="castellano")
        ruta_xlsx = out / f"Datos_filtrados_{nombre}.xlsx"
        with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as writer:
            df_a.to_excel(writer, sheet_name="Datos_filtrados", index=False)
            tabla_p_es.to_excel(writer, sheet_name="Valores_por_partido", index=False)
            resumen_global_es.to_excel(writer, sheet_name="Resumen_cientifico", index=False)
            if not tabla_dificultad_es.empty:
                tabla_dificultad_es.to_excel(writer, sheet_name="Dificultad_comp", index=False)
            if not tabla_criticos_es.empty:
                tabla_criticos_es.to_excel(writer, sheet_name="Partidos_criticos", index=False)
            pd.DataFrame({"Hallazgos": hallazgos_es}).to_excel(writer, sheet_name="Hallazgos", index=False)
            pd.DataFrame({"Discusion_integrada": [discusion_es]}).to_excel(writer, sheet_name="Discusion", index=False)
            for titulo, tabla in resumenes.items():
                tabla.to_excel(writer, sheet_name=nombre_seguro(titulo)[:31], index=False)

        def construir_figuras(idioma_actual):
            figs_id = out / ("Figuras_EN" if idioma_actual == "ingles" else "Figuras_ES")
            figs_id.mkdir(exist_ok=True)
            figuras_locales = []

            # Figuras analÃ­ticas por partido solo para variables clave, evitando repeticiÃ³n excesiva.
            claves_preferentes = ["IGFA ponderado POST", "IGFA simple POST", "RPE POST", "Molestias (MS) POST", "Capacidad fÃ­sica (PF) POST", "Capacidad mental (MF) POST", "Confianza (RC) POST", "Î”IGFA ponderado", "Î”IGFA simple", "Î”RPE", "Î”MS", "Î”PF", "Î”MF", "Î”RC"]
            etiquetas_a_graficar = [e for e in claves_preferentes if e in variables_analisis]
            if not etiquetas_a_graficar:
                etiquetas_a_graficar = list(variables_analisis.keys())[:8]

            if "partidos" in tipos:
                for etiqueta in etiquetas_a_graficar:
                    col = variables_analisis[etiqueta]
                    ruta_png = figs_id / f"{nombre_seguro(etiqueta)}_por_partidos_{idioma_actual}.png"
                    stats = grafico_variable_por_partidos(df_a, cols["partido"], etiqueta, col, ruta_png, idioma=idioma_actual)
                    if stats:
                        figuras_locales.append({
                            "titulo": f"{etiqueta_idioma(etiqueta, idioma_actual)} {TEXTOS[idioma_actual]['por_partidos']}",
                            "ruta": str(ruta_png),
                            "interpretacion": interpretar_partido(etiqueta, stats, idioma=idioma_actual)
                        })

            if "competitiva" in tipos and "Resumen por categorÃ­a competitiva" in resumenes:
                for etiqueta in etiquetas_a_graficar[:8]:
                    titulo = f"{etiqueta_idioma(etiqueta, idioma_actual)} {TEXTOS[idioma_actual]['por_competitiva']}"
                    ruta_png = figs_id / f"{nombre_seguro(etiqueta)}_por_categoria_competitiva_{idioma_actual}.png"
                    stats = grafico_variable_por_grupo(resumenes["Resumen por categorÃ­a competitiva"], "CategorÃ­a competitiva", etiqueta, ruta_png, titulo, idioma=idioma_actual)
                    if stats:
                        figuras_locales.append({
                            "titulo": titulo,
                            "ruta": str(ruta_png),
                            "interpretacion": interpretar_grupo(etiqueta, stats, "por categorÃ­a competitiva", idioma=idioma_actual)
                        })

            ruta_heatmap = figs_id / f"heatmap_partido_variables_{idioma_actual}.png"
            stats_hm = grafico_heatmap_partido_variables(df_a, cols["partido"], variables_analisis, ruta_heatmap, idioma=idioma_actual)
            if stats_hm:
                figuras_locales.append({
                    "titulo": "Heatmap Partido Ã— Variables" if idioma_actual != "ingles" else "Match Ã— Variables heatmap",
                    "ruta": str(ruta_heatmap),
                    "interpretacion": "El mapa de calor resume la carga relativa por partido y variable. En las variables positivas se invierte la escala para que los tonos de mayor carga representen situaciones relativamente menos favorables."
                })
            return figuras_locales

        filtros = {
            "Partidos": seleccion_partidos if seleccion_partidos else ["Todos"],
            "CategorÃ­a competitiva": seleccion_comp if seleccion_comp else ["Todas/no aplica"],
            "Tipo de anÃ¡lisis": sorted(list(tipos)),
            "Ãndices incluidos": sorted(list(indices_variables.keys())) if indices_variables else ["No incluidos"],
            "Excel": ruta_excel.name,
            "Plantilla": ruta_plantilla.name,
        }

        for idioma in idiomas:
            sufijo_idioma = "EN" if idioma == "ingles" else "ES"
            ruta_docx = out / f"Informe_{nombre}_{sufijo_idioma}.docx"
            tabla_p_idioma = tabla_partidos(df_a, cols["partido"], variables, indices_variables, idioma=idioma)
            hallazgos = hallazgos_principales(df_a, cols["partido"], variables_base, variables_analisis, cols.get("competitiva"), idioma=idioma)
            crear_informe(
                ruta_plantilla=ruta_plantilla,
                ruta_salida=ruta_docx,
                datos_arbitro=datos_arbitro,
                ruta_foto=foto,
                filtros=filtros,
                tabla_p=tabla_p_idioma,
                resumenes=resumenes,
                figuras=construir_figuras(idioma),
                tipos=tipos,
                n_obs=len(df_a),
                resumen_global=tabla_resumen_global(df_a, cols["partido"], variables_analisis, cols.get("competitiva"), idioma=idioma),
                hallazgos=hallazgos,
                interpretacion=None,
                conclusion_auto=None,
                tabla_dificultad=tabla_comparativa_dificultad(df_a, cols.get("competitiva"), variables_analisis, idioma=idioma),
                tabla_criticos=discutir_partidos_criticos(df_a, cols["partido"], variables_analisis, cols.get("competitiva"), idioma=idioma),
                discusion_integrada=discusion_integrada_una_pagina(df_a, cols["partido"], variables_analisis, cols.get("competitiva"), hallazgos, idioma=idioma),
                idioma=idioma,
            )
            print(f"Informe Word generado: {ruta_docx}")
            if "pdf" in formatos:
                ruta_pdf = ruta_docx.with_suffix(".pdf")
                if PDF_OK:
                    try:
                        docx2pdf_convert(str(ruta_docx), str(ruta_pdf))
                        print(f"Informe PDF generado: {ruta_pdf}")
                    except Exception as e:
                        print(f"No se pudo generar PDF automÃ¡ticamente: {e}")
                else:
                    print("PDF no disponible. Instala con: pip install docx2pdf")
            if formatos == ["pdf"] and ruta_docx.exists():
                try:
                    ruta_docx.unlink()
                except Exception:
                    pass
        print(f"Excel generado:   {ruta_xlsx}")

    print("\nPROCESO FINALIZADO.")
    print(f"Resultados guardados en: {resultados}")

if __name__ == "__main__":
    main()


