from pathlib import Path
import pandas as pd

ORIGEN = Path("Datos/Datos.xlsx")
SALIDA = Path("examples/anonymized_data/Datos_example.xlsx")

SALIDA.parent.mkdir(parents=True, exist_ok=True)

if not ORIGEN.exists():
    raise FileNotFoundError(f"No se encontró: {ORIGEN.resolve()}")

hojas = pd.read_excel(ORIGEN, sheet_name=None)

principal = hojas["Partidos_Cuestionarios"].copy()

# Seleccionar como máximo 3 árbitros y 4 partidos por árbitro.
arbitros_originales = (
    principal["ref_id"]
    .dropna()
    .astype(str)
    .drop_duplicates()
    .head(3)
    .tolist()
)

principal = (
    principal[principal["ref_id"].astype(str).isin(arbitros_originales)]
    .groupby("ref_id", group_keys=False)
    .head(4)
    .copy()
)

partidos_originales = (
    principal["match_id"]
    .dropna()
    .astype(str)
    .drop_duplicates()
    .tolist()
)

mapa_arbitros = {
    original: f"REF_{i:03d}"
    for i, original in enumerate(arbitros_originales, 1)
}

mapa_nombres = {
    original: f"Referee {i:03d}"
    for i, original in enumerate(arbitros_originales, 1)
}

mapa_partidos = {
    original: f"MATCH_{i:03d}"
    for i, original in enumerate(partidos_originales, 1)
}

equipos = []
for columna in ["team_a", "team_b"]:
    if columna in principal.columns:
        equipos.extend(
            principal[columna].dropna().astype(str).drop_duplicates().tolist()
        )

equipos = list(dict.fromkeys(equipos))

mapa_equipos = {
    original: f"TEAM_{i:03d}"
    for i, original in enumerate(equipos, 1)
}


def anonimizar_comun(df):
    df = df.copy()

    if "ref_id" in df.columns:
        df = df[df["ref_id"].astype(str).isin(arbitros_originales)]
        df["ref_id"] = df["ref_id"].astype(str).map(mapa_arbitros)

    if "questionnaire_ref_id" in df.columns:
        df["questionnaire_ref_id"] = (
            df["questionnaire_ref_id"].astype(str).map(mapa_arbitros)
        )

    if "questionnaire_referee_name" in df.columns:
        df["questionnaire_referee_name"] = (
            df["questionnaire_ref_id"]
            .map({
                nuevo: f"Referee {i:03d}"
                for i, nuevo in enumerate(mapa_arbitros.values(), 1)
            })
        )

    if "match_id" in df.columns:
        df = df[df["match_id"].astype(str).isin(partidos_originales)]
        df["match_id"] = df["match_id"].astype(str).map(mapa_partidos)

    if "team_a" in df.columns:
        df["team_a"] = df["team_a"].astype(str).map(mapa_equipos)

    if "team_b" in df.columns:
        df["team_b"] = df["team_b"].astype(str).map(mapa_equipos)

    if "date" in df.columns:
        fechas = {
            partido: pd.Timestamp("2026-01-01") + pd.Timedelta(days=i * 2)
            for i, partido in enumerate(mapa_partidos.values())
        }

        if "match_id" in df.columns:
            df["date"] = df["match_id"].map(fechas)

    return df


salida = {}

# REFEREES: crear perfiles completamente ficticios.
referees = hojas["REFEREES"].copy()
referees = referees[
    referees["REF_ID"].astype(str).isin(arbitros_originales)
].copy()

referees["REF_ID"] = referees["REF_ID"].astype(str).map(mapa_arbitros)
referees["NAME"] = [
    f"Referee {i:03d}" for i in range(1, len(referees) + 1)
]
referees["SUBJECT"] = [
    f"SUBJECT_{i:03d}" for i in range(1, len(referees) + 1)
]
referees["NATIONALITY"] = "Anonymized"
referees["AGE"] = [30 + i for i in range(len(referees))]
referees["SEX"] = "X"
referees["HEIGHT"] = [175 + i for i in range(len(referees))]
referees["WEIGHT"] = [72 + i for i in range(len(referees))]
referees["EXPERIENCE"] = [5 + i for i in range(len(referees))]
referees["LEVEL"] = "Example"
referees["mail"] = [
    f"referee{i:03d}@example.org"
    for i in range(1, len(referees) + 1)
]

salida["REFEREES"] = referees
salida["Partidos_Cuestionarios"] = anonimizar_comun(
    hojas["Partidos_Cuestionarios"]
)

# Hojas relacionadas.
for nombre in [
    "Partidos",
    "Cuestionarios",
    "Partidos_Nombre_Arbitro",
    "Control_Cruce_Dificultad",
]:
    if nombre in hojas:
        salida[nombre] = anonimizar_comun(hojas[nombre])

# Hojas sin información personal directa.
for nombre in ["Diccionario_Variables", "¿preguntas"]:
    if nombre in hojas:
        salida[nombre] = hojas[nombre].copy()

with pd.ExcelWriter(SALIDA, engine="openpyxl") as writer:
    for nombre, df in salida.items():
        df.to_excel(writer, sheet_name=nombre[:31], index=False)

print(f"Archivo anonimizado creado: {SALIDA.resolve()}")
print(f"Árbitros ficticios: {len(referees)}")
print(f"Partidos ficticios: {len(partidos_originales)}")
