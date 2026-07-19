"""Language and translation utilities."""

import pandas as pd

from r7rfa.constants import TRAD_EN

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
