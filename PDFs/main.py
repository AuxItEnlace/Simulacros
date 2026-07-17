import csv
import gc
import os
import time
import pandas as pd
import threading
import keyboard
import win32com.client as win32
from tqdm import tqdm

# === CONFIGURACIÓN ===
plantilla = os.path.abspath("config/plantilla.docx")
datos = os.path.abspath("data/data.csv")
salida_pdf = os.path.abspath("pdfs_por_grado")
salida_temp_excel = os.path.abspath("temp_excels")
os.makedirs(salida_pdf, exist_ok=True)
os.makedirs(salida_temp_excel, exist_ok=True)

# === CANCELACIÓN ===
cancelado = False
def detectar_cancelacion():
    global cancelado
    tqdm.write("⏳ Presiona ALT + Q para cancelar en cualquier momento.")
    keyboard.wait("alt+q")
    cancelado = True
    tqdm.write("❌ Cancelación solicitada por el usuario.")

# === CARGAR DATOS ===
df = pd.read_csv(datos, encoding='utf-8-sig', sep=';')
df.columns = df.columns.str.strip().str.upper()
df["GRADO"] = df["GRADO"].astype(str).str.strip().str.upper()

if "CALIFICACIÓN" in df.columns:
    df["CALIFICACIÓN"] = df["CALIFICACIÓN"].astype(str).str.strip()

grados = sorted(df["GRADO"].dropna().unique(), key=lambda x: int(''.join(c for c in x if c.isdigit()) or '0'))


def save_csv_with_text_format(df_to_save, csv_path, text_columns=None):
    if text_columns is None:
        text_columns = []

    fieldnames = list(df_to_save.columns)
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for _, row in df_to_save.iterrows():
            row_data = row.to_dict()
            for col in text_columns:
                if col in row_data and row_data[col] is not None:
                    row_data[col] = str(row_data[col])
            writer.writerow(row_data)


def _open_merge_and_export(word, template_path, data_excel_path, pdf_output_path):
    plantilla_doc = word.Documents.Open(template_path)
    plantilla_doc.MailMerge.OpenDataSource(
        Name=data_excel_path,
        Connection=(
            'Provider=Microsoft.ACE.OLEDB.12.0;'
            'Data Source=' + data_excel_path + ';'
            'Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0";'
        ),
        SQLStatement='SELECT * FROM [Hoja1$]',
    )
    plantilla_doc.MailMerge.Destination = 0
    plantilla_doc.MailMerge.Execute(Pause=False)

    merged_doc = word.ActiveDocument
    merged_doc.ExportAsFixedFormat(pdf_output_path, 17)
    merged_doc.Close(False)
    plantilla_doc.Close(False)
    del merged_doc, plantilla_doc
    gc.collect()


# === INICIAR WORD ===
start_time = time.time()
word = None

try:
    threading.Thread(target=detectar_cancelacion, daemon=True).start()

    tqdm.write("🚀 Iniciando Word...")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    total_pdfs_grupo = len(grados)
    total_pdfs_individual = len(df)
    total_pdfs = total_pdfs_grupo + total_pdfs_individual
    pdfs_generados = 0

    barra_principal = tqdm(total=total_pdfs, desc="📊 Avance total", unit="PDF",
                           bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]",
                           position=0, leave=True)

    for grado in grados:
        if cancelado:
            raise KeyboardInterrupt("Cancelado por el usuario.")

        df_grado = df[df["GRADO"] == grado]
        if df_grado.empty:
            tqdm.write(f"⚠️  Grado {grado} está vacío. Se omite.")
            continue

        # === PDF GRUPAL POR GRADO ===
        grado_csv = os.path.join(salida_temp_excel, f"{grado}_GRUPO.csv")
        save_csv_with_text_format(df_grado, grado_csv, text_columns=["CALIFICACIÓN"])

        pdf_path_grupo = os.path.join(salida_pdf, f"{grado}.pdf")
        barra_principal.set_description(f"📤 {grado} → grupal")
        _open_merge_and_export(word, plantilla, grado_csv, pdf_path_grupo)

        pdfs_generados += 1
        barra_principal.update(1)
        barra_principal.set_postfix_str(f"Grado {grado} | Grupal ✓")

        # === PDF INDIVIDUAL POR ESTUDIANTE ===
        carpeta_individual = os.path.join(salida_pdf, str(grado))
        os.makedirs(carpeta_individual, exist_ok=True)

        estudiantes = list(df_grado.iterrows())
        for idx, (_, fila) in enumerate(estudiantes):
            if cancelado:
                raise KeyboardInterrupt("Cancelado por el usuario.")

            nombre_estudiante = fila["NOMBRE"].strip().replace("/", "-").replace("\\", "-")
            nombre_archivo = os.path.join(salida_temp_excel, f"{grado}_{nombre_estudiante}.csv")
            save_csv_with_text_format(fila.to_frame().T, nombre_archivo, text_columns=["CALIFICACIÓN"])

            pdf_individual = os.path.join(carpeta_individual, f"{nombre_estudiante}.pdf")
            barra_principal.set_description(f"📄 {grado} → {nombre_estudiante}")
            _open_merge_and_export(word, plantilla, nombre_archivo, pdf_individual)

            pdfs_generados += 1
            barra_principal.update(1)
            barra_principal.set_postfix_str(f"Grado {grado} | {idx+1}/{len(estudiantes)}")

    barra_principal.set_description("✅ Completado")
    barra_principal.close()

except KeyboardInterrupt as e:
    tqdm.write(f"\n🛑 {e}")
except Exception as e:
    tqdm.write(f"\n❌ Error inesperado: {e}")
finally:
    tqdm.write("🧹 Cerrando Word...")
    if word is not None:
        try:
            word.Quit()
        except Exception as close_error:
            tqdm.write(f"⚠️  Word no pudo cerrarse automáticamente: {close_error}")
    del word
    gc.collect()
    tiempo_total = time.time() - start_time
    tqdm.write(f"✅ Finalizado en {tiempo_total:.2f} segundos.")
