import re
import statistics
from pathlib import Path

from openpyxl import Workbook

from config import load_config
from data_processor import load_and_sort_data, calculate_columns
from excel_writer import (
    create_base_sheets,
    write_dataframe_to_sheet,
    build_base_header,
    build_detailed_header,
    process_grade_sheets,
    add_summary_by_group,
    format_percentages,
    save_workbook,
)
from excel_formatter import format_workbook


def main() -> None:
    config = load_config()

    base_dir = Path(__file__).resolve().parent
    csv_path = base_dir / "data" / "data.csv"
    data = load_and_sort_data(str(csv_path))
    data = calculate_columns(data)

    if config.get("Calificacion", "FALSE").upper() == "TRUE":
        data.insert(0, "CALIFICACION", round(data["PUNTAJE"] * 5 / 100, 1))

    sede = config.get("Sede", "")
    school_tag = sede[:10] if len(sede) > 10 else sede
    if not school_tag:
        school_tag = config["SchoolName"][:10]

    # Determine which grades have data (extract numeric prefix from GRADO)
    data["_GRADO_NUM"] = data["GRADO"].str.extract(r"^(\d+)").astype(int)
    active_grades = sorted(data["_GRADO_NUM"].unique())
    data.drop(columns=["_GRADO_NUM"], inplace=True)

    sheet_names = [
        "Reporte General",
        "Detallado Grados",
        f"BD-RESULTADOS-{school_tag}",
    ] + [f"GRADO ({i})" for i in active_grades]

    wb = create_base_sheets(Workbook(), sheet_names)

    # Raw data sheet
    write_dataframe_to_sheet(wb, sheet_names[2], data)

    # Write headers for summary sheets
    wb[sheet_names[0]].append(build_base_header())
    wb[sheet_names[1]].append(build_detailed_header())

    # Per-grade averages (only sheets that exist)
    avg = process_grade_sheets(wb, data, sheet_names[3:])

    avg_sorted = sorted(
        avg, key=lambda x: int(re.match(r"(\d+)", str(x[0])).group(1))
    )

    # Detailed summary (sheet 1)
    school_full = f"{config['SchoolName']} {sede}".strip()
    grade_summaries = add_summary_by_group(
        wb[sheet_names[1]], avg_sorted, school_full
    )

    # Compute overall school std dev from grade summaries (column index 3 = % CORRECTAS)
    pct_values = [row[3] for row in grade_summaries]
    if len(pct_values) > 1:
        school_std = round(statistics.stdev(pct_values), 2)
    else:
        school_std = 0

    for row in grade_summaries:
        wb[sheet_names[0]].append(row)

    wb[sheet_names[0]].append(["Desviación Estándar del Colegio", school_std])

    # Apply percentage formatting (converts "85%" → 0.85)
    format_percentages(wb, sheet_names)

    # Apply visual styles (colors, borders, two‑row headers, banding)
    grade_worksheets = [wb[name] for name in sheet_names[3:]]
    format_workbook(
        wb[sheet_names[0]],
        wb[sheet_names[1]],
        wb[sheet_names[2]],
        grade_worksheets,
    )

    path = save_workbook(wb, config, str(base_dir))
    print(f"Archivo guardado: {path}")


if __name__ == "__main__":
    main()
