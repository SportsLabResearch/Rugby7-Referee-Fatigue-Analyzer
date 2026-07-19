"""
Core scientific analysis routines.

The functions in this module calculate fatigue indices and produce
descriptive summaries from referee observations.
"""

import numpy as np
import pandas as pd

from r7rfa.constants import *
from r7rfa.language import *
from r7rfa.utils import *

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
        nuevos["ÃŽâ€IGFA simple"] = col

    if "IGFA ponderado PRE" in nuevos and "IGFA ponderado POST" in nuevos:
        col = "Delta_IGFA_ponderado"
        df[col] = (df[nuevos["IGFA ponderado POST"]] - df[nuevos["IGFA ponderado PRE"]]).round(2)
        nuevos["ÃŽâ€IGFA ponderado"] = col

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
