import pandas as pd
from typing import List


def load_and_sort_data(filepath: str) -> pd.DataFrame:
    data = pd.read_csv(filepath, encoding="utf-8", engine="python")
    if data.empty:
        raise ValueError("El archivo CSV está vacío.")

    data["GRADO"] = data["GRADO"].astype(str)
    data[["GRADO_NUM", "GRADO_LETRA"]] = data["GRADO"].str.extract(r"(\d+)([A-Z])?")
    data["GRADO_NUM"] = data["GRADO_NUM"].astype(int)
    data["% CORRECTAS"] = (
        data["% CORRECTAS"].astype(str).str.replace("%", "").astype(float)
    )

    data.sort_values(
        by=["GRADO_NUM", "GRADO_LETRA", "% CORRECTAS", "NOMBRE"],
        ascending=[True, True, False, True],
        inplace=True,
    )

    data.drop(columns=["GRADO_NUM", "GRADO_LETRA"], inplace=True)

    return data


def calculate_columns(data: pd.DataFrame) -> pd.DataFrame:
    data["PUNTAJE_PONDERADO"] = (
        data["Correctas C1"] * 1
        + data["Correctas C2"] * 2
        + data["Correctas C3"] * 3
        + data["CORRECTAS"] * 4
    )

    data.insert(
        0,
        "PUESTO",
        data.groupby("GRADO")["PUNTAJE_PONDERADO"]
        .rank(ascending=False)
        .astype(int),
    )

    data.drop(columns=["PUNTAJE_PONDERADO"], inplace=True)

    data["GRADO"] = data["GRADO"].astype(str)
    data[["GRADO_NUM", "GRADO_LETRA"]] = data["GRADO"].str.extract(r"(\d+)([A-Z])?")
    data["GRADO_NUM"] = data["GRADO_NUM"].astype(int)

    data.sort_values(
        by=["GRADO_NUM", "GRADO_LETRA", "PUESTO"],
        ascending=[True, True, True],
        inplace=True,
    )

    data.drop(columns=["GRADO_NUM", "GRADO_LETRA"], inplace=True)

    data.insert(0, "CONTAR", data.groupby("GRADO").cumcount(ascending=False) + 1)
    data.insert(
        1,
        "PERCENTIL",
        data.groupby("GRADO")["% CORRECTAS"].transform(
            lambda x: (x.rank(pct=True) * 100).astype(int)
        ),
    )
    data.insert(2, "PUNTAJE", data["% CORRECTAS"].round(2))

    data["% CORRECTAS"] = (data["% CORRECTAS"].astype(int).astype(str) + "%")

    return data


def calculate_percentiles(scores: List[float]) -> List[int]:
    n = len(scores)
    sorted_scores = sorted(scores)
    percentiles: List[int] = []
    for p in scores:
        positions = [i + 1 for i, val in enumerate(sorted_scores) if val == p]
        avg_rank = sum(positions) / len(positions)
        percentiles.append(round((avg_rank / n) * 100))
    return percentiles
