"""Chronological ordering utilities for match observations."""

import pandas as pd

from r7rfa.utils import detectar_columna


def ordenar_cronologicamente(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ordena las observaciones por fecha y hora cuando esas columnas existen.

    Si no se detecta una fecha válida, conserva el orden original.
    """
    if df is None or df.empty:
        return df

    resultado = df.copy()

    col_fecha = detectar_columna(
        resultado,
        ["date", "fecha", "match_date", "fecha_partido"],
        posicion=None,
    )

    col_hora = detectar_columna(
        resultado,
        ["time", "hora", "match_time", "hora_partido"],
        posicion=None,
    )

    if col_fecha is None:
        return resultado.reset_index(drop=True)

    fecha = pd.to_datetime(
        resultado[col_fecha],
        errors="coerce",
        dayfirst=True,
    )

    if col_hora is not None:
        hora_texto = resultado[col_hora].astype(str).str.strip()

        fecha_hora = pd.to_datetime(
            fecha.dt.strftime("%Y-%m-%d") + " " + hora_texto,
            errors="coerce",
        )

        resultado["_ORDEN_CRONOLOGICO"] = fecha_hora.fillna(fecha)
    else:
        resultado["_ORDEN_CRONOLOGICO"] = fecha

    resultado["_ORDEN_ORIGINAL"] = range(len(resultado))

    resultado = (
        resultado
        .sort_values(
            ["_ORDEN_CRONOLOGICO", "_ORDEN_ORIGINAL"],
            kind="stable",
            na_position="last",
        )
        .drop(columns=["_ORDEN_CRONOLOGICO", "_ORDEN_ORIGINAL"])
        .reset_index(drop=True)
    )

    return resultado
