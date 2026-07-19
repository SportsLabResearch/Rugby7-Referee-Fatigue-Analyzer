"""Tests for chronological ordering utilities."""

import pandas as pd

from r7rfa.chronology import ordenar_cronologicamente


def test_ordenar_por_fecha_y_hora():
    datos = pd.DataFrame(
        {
            "Fecha": [
                "03/06/2025",
                "01/06/2025",
                "01/06/2025",
                "02/06/2025",
            ],
            "Hora": [
                "10:00",
                "16:00",
                "09:00",
                "12:00",
            ],
            "Partido": ["D", "B", "A", "C"],
        }
    )

    resultado = ordenar_cronologicamente(datos)

    assert resultado["Partido"].tolist() == ["A", "B", "C", "D"]


def test_conservar_orden_si_no_hay_fecha():
    datos = pd.DataFrame(
        {
            "Partido": ["C", "A", "B"],
            "RPE": [5, 3, 4],
        }
    )

    resultado = ordenar_cronologicamente(datos)

    assert resultado["Partido"].tolist() == ["C", "A", "B"]
    assert resultado.index.tolist() == [0, 1, 2]


def test_fechas_invalidas_quedan_al_final():
    datos = pd.DataFrame(
        {
            "Fecha": ["fecha desconocida", "02/06/2025", "01/06/2025"],
            "Hora": ["", "10:00", "09:00"],
            "Partido": ["X", "B", "A"],
        }
    )

    resultado = ordenar_cronologicamente(datos)

    assert resultado["Partido"].tolist() == ["A", "B", "X"]


def test_dataframe_vacio():
    datos = pd.DataFrame()

    resultado = ordenar_cronologicamente(datos)

    assert resultado.empty


def test_no_modifica_dataframe_original():
    datos = pd.DataFrame(
        {
            "Fecha": ["02/06/2025", "01/06/2025"],
            "Partido": ["B", "A"],
        }
    )

    original = datos.copy(deep=True)

    ordenar_cronologicamente(datos)

    pd.testing.assert_frame_equal(datos, original)
