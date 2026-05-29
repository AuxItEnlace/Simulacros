import re
import statistics

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
    GRADE_RANGE,
)


def main() -> None:
    config = load_config()

    data = load_and_sort_data("data/data.csv")
    data = calculate_columns(data)

    school_name = config["Sede"][:10] if len(config["Sede"]) > 10 else config["Sede"]

    sheet_names = [
        "Reporte General",
        "Detallado Grados",
        f"BD-RESULTADOS-{school_name}",
    ] + [f"GRADO ({i})" for i in range(GRADE_RANGE)]

    wb = create_base_sheets(Workbook(), sheet_names)

    # Raw data sheet
    write_dataframe_to_sheet(wb, sheet_names[2], data)

    # Write headers for summary sheets
    wb[sheet_names[0]].append(build_base_header())
    wb[sheet_names[1]].append(build_detailed_header())

    # Per-grade averages
    avg = process_grade_sheets(wb, data, sheet_names[3:])

    avg_sorted = sorted(
        avg, key=lambda x: int(re.match(r"(\d+)", str(x[0])).group(1))
    )

    # Detailed summary (sheet 1)
    school_full = f"{config['SchoolName']} {config['Sede']}"
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

    # Apply percentage formatting
    format_percentages(wb, sheet_names)

    path = save_workbook(wb, config)
    print(f"Archivo guardado: {path}")


if __name__ == "__main__":
    main()
