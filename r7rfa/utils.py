"""Generic text and DataFrame utilities."""

import re
import textwrap
import unicodedata

import pandas as pd

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

def valores_unicos(df, col):
    if col is None or col not in df.columns:
        return []
    vals = [limpiar_texto(v) for v in df[col].dropna().unique()]
    vals = [v for v in vals if v]
    return sorted(vals, key=lambda x: normalizar(x))

def filtrar(df, col, seleccion):
    if col is None or not seleccion:
        return df.copy()
    sel = {normalizar(v) for v in seleccion}
    return df[df[col].map(lambda x: normalizar(x) in sel)].copy()
