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
    print("⏳ Presiona ALT + Q para cancelar en cualquier momento.")
    keyboard.wait("alt+q")
    cancelado = True
    print("❌ Cancelación solicitada por el usuario.")

# === CARGAR DATOS ===
df = pd.read_csv(datos, encoding='utf-8-sig', sep=';')
df.columns = df.columns.str.strip().str.upper()
df["GRADO"] = df["GRADO"].astype(str).str.strip().str.upper()

if "CALIFICACIÓN" in df.columns:
    df["CALIFICACIÓN"] = (
        df["CALIFICACIÓN"].astype(str).str.strip().str.replace("(", "").str.replace(")", "").str.replace(",", ".").astype(float).map(lambda x: f"{x:.1f}")
    )

grados = sorted(df["GRADO"].dropna().unique(), key=lambda x: int(''.join(c for c in x if c.isdigit()) or '0'))


def _open_merge_and_export(word, template_path, data_excel_path, pdf_output_path):
    plantilla_doc = word.Documents.Open(template_path)
    plantilla_doc.MailMerge.OpenDataSource(
        Name=data_excel_path,
        Connection=(
            'Provider=Microsoft.ACE.OLEDB.12.0;'
            'Data Source=' + data_excel_path + ';'
            'Extended Properties="Excel 12.0 Xml;HDR=YES";'
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

    print("🚀 Iniciando Word...")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = False

    for grado in tqdm(grados, desc="📚 Procesando grados", unit="grado"):
        if cancelado:
            raise KeyboardInterrupt("Cancelado por el usuario.")

        df_grado = df[df["GRADO"] == grado]
        if df_grado.empty:
            print(f"⚠️ Grado {grado} está vacío. Se omite.")
            continue

        # === PDF GRUPAL POR GRADO ===
        print(f"\n🔄 Procesando grado completo: {grado}")
        grado_excel = os.path.join(salida_temp_excel, f"{grado}_GRUPO.xlsx")
        df_grado.to_excel(grado_excel, index=False, sheet_name='Hoja1')

        pdf_path_grupo = os.path.join(salida_pdf, f"{grado}.pdf")
        print(f"📤 Exportando PDF grupal: {pdf_path_grupo}")
        _open_merge_and_export(word, plantilla, grado_excel, pdf_path_grupo)

        # === PDF INDIVIDUAL POR ESTUDIANTE ===
        print(f"👤 Generando PDFs individuales para {grado}...")
        carpeta_individual = os.path.join(salida_pdf, str(grado))
        os.makedirs(carpeta_individual, exist_ok=True)

        for _, fila in df_grado.iterrows():
            if cancelado:
                raise KeyboardInterrupt("Cancelado por el usuario.")

            nombre_estudiante = fila["NOMBRE"].strip().replace("/", "-").replace("\\", "-")
            nombre_archivo = os.path.join(salida_temp_excel, f"{grado}_{nombre_estudiante}.xlsx")
            fila.to_frame().T.to_excel(nombre_archivo, index=False, sheet_name="Hoja1")

            pdf_individual = os.path.join(carpeta_individual, f"{nombre_estudiante}.pdf")
            print(f"📄 Guardando PDF individual: {pdf_individual}")
            _open_merge_and_export(word, plantilla, nombre_archivo, pdf_individual)


except KeyboardInterrupt as e:
    print(f"\n🛑 {e}")
except Exception as e:
    print(f"\n❌ Error inesperado: {e}")
finally:
    print("🧹 Cerrando Word...")
    if word is not None:
        try:
            word.Quit()
        except Exception as close_error:
            print(f"⚠️ Word no pudo cerrarse automáticamente: {close_error}")
    del word
    gc.collect()
    tiempo_total = time.time() - start_time
    print(f"✅ Finalizado en {tiempo_total:.2f} segundos.")
