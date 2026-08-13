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
df = df.dropna(how="all")
df.columns = df.columns.str.strip().str.upper()
df["GRADO"] = df["GRADO"].astype(str).str.strip().str.upper()
if "GRUPO" in df.columns:
    df["GRUPO"] = df["GRUPO"].astype(str).str.strip().str.upper()
    df["GRUPO_COMPLETO"] = df["GRADO"] + df["GRUPO"]
else:
    df["GRUPO_COMPLETO"] = df["GRADO"]

if "CALIFICACIÓN" in df.columns:
    df["CALIFICACIÓN"] = df["CALIFICACIÓN"].astype(str).str.strip()

if "PUESTO (TCT)" in df.columns:
    df["PUESTO (TCT)"] = pd.to_numeric(df["PUESTO (TCT)"], errors="coerce")

df["GRADO_NUM"] = df["GRADO"].str.extract(r"(\d+)").astype(int)
columna_grupo = "GRUPO" if "GRUPO" in df.columns else "GRUPO_COMPLETO"
columnas_orden = ["GRADO_NUM", columna_grupo]
ascending_orden = [True, True]
if "PUESTO (TCT)" in df.columns:
    columnas_orden.append("PUESTO (TCT)")
    ascending_orden.append(True)
df.sort_values(
    by=columnas_orden,
    ascending=ascending_orden,
    inplace=True,
)
df.drop(columns=["GRADO_NUM"], inplace=True)


def _grado_key(x):
    numero = int(''.join(c for c in x if c.isdigit()) or '0')
    letra = ''.join(c for c in x if not c.isdigit())
    return (numero, letra)


grados = sorted(df["GRUPO_COMPLETO"].dropna().unique(), key=_grado_key)


def save_csv_with_text_format(df_to_save, excel_path, text_columns=None):
    if text_columns is None:
        text_columns = []

    df_out = df_to_save.copy()
    for col in text_columns:
        if col in df_out.columns:
            df_out[col] = df_out[col].astype(str)

    df_out.to_excel(excel_path, index=False, engine='openpyxl', sheet_name='Hoja1')


def _open_merge_and_export(word, template_path, data_excel_path, pdf_output_path):
    plantilla_doc = word.Documents.Open(template_path)
    plantilla_doc.MailMerge.OpenDataSource(
        Name=data_excel_path,
        AddToRecentFiles=False,
        Revert=False,
        Format=0,
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
    word.Visible = True
    word.DisplayAlerts = False
    word.Options.ConfirmConversions = False

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

        df_grado = df[df["GRUPO_COMPLETO"] == grado]
        if df_grado.empty:
            tqdm.write(f"⚠️  Grado {grado} está vacío. Se omite.")
            continue

        # === PDF GRUPAL POR GRADO ===
        grado_csv = os.path.join(salida_temp_excel, f"{grado}_GRUPO.xlsx")
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

            nombre_estudiante = fila["NOMBRE COMPLETO"].strip().replace("/", "-").replace("\\", "-")
            nombre_archivo = os.path.join(salida_temp_excel, f"{grado}_{nombre_estudiante}.xlsx")
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
