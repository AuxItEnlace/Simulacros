import win32com.client as win32
import os
import time
import pandas as pd
import threading
import keyboard
from tqdm import tqdm

# === CONFIGURACIÓN ===
plantilla = os.path.abspath("data/plantilla.docx")
datos = os.path.abspath("data/data.csv")
salida_pdf = os.path.abspath("pdfs_por_grado")
salida_docx = os.path.abspath("temp_docs_por_grado")
salida_temp_excel = os.path.abspath("temp_excels")
os.makedirs(salida_pdf, exist_ok=True)
os.makedirs(salida_docx, exist_ok=True)
os.makedirs(salida_temp_excel, exist_ok=True)

# === CANCELACIÓN ===
cancelado = False
def detectar_cancelacion():
    global canclselado
    print("⏳ Presiona ALT + Q para cancelar en cualquier momento.")
    keyboard.wait("alt+q")
    cancelado = True
    print("❌ Cancelación solicitada por el usuario.")

# === CARGAR DATOS ===
df = pd.read_csv(datos, encoding='utf-8-sig', sep=';')
df.columns = df.columns.str.strip().str.upper()
grados = df["GRADO"].dropna().unique()
grados.sort()

# === INICIAR WORD ===
print("🚀 Iniciando Word...")
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = True

start_time = time.time()

try:
    threading.Thread(target=detectar_cancelacion, daemon=True).start()

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

        plantilla_doc = word.Documents.Open(plantilla)
        plantilla_doc.MailMerge.OpenDataSource(
            Name=grado_excel,
            ConfirmConversions=False,
            ReadOnly=True,
            LinkToSource=True,
            AddToRecentFiles=False,
            Revert=False,
            Format=0,
            Connection='Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + grado_excel + ';Extended Properties="Excel 12.0 XML;HDR=YES";',
            SQLStatement='SELECT * FROM [Hoja1$]'
        )

        plantilla_doc.MailMerge.Destination = 0  # nuevo documento
        plantilla_doc.MailMerge.Execute(Pause=False)
        plantilla_doc.Close(False)

        combined_doc = word.ActiveDocument
        pdf_path_grupo = os.path.join(salida_pdf, f"{str(grado)}.pdf")
        print(f"📤 Exportando PDF grupal: {pdf_path_grupo}")
        combined_doc.ExportAsFixedFormat(pdf_path_grupo, 17)
        combined_doc.Close(False)

        # === PDF INDIVIDUAL POR ESTUDIANTE ===
        print(f"👤 Generando PDFs individuales para {grado}...")
        carpeta_individual = os.path.join(salida_pdf, str(grado))
        os.makedirs(carpeta_individual, exist_ok=True)

        for _, fila in df_grado.iterrows():
            if cancelado:
                raise KeyboardInterrupt("Cancelado por el usuario.")

            nombre_estudiante = fila["NOMBRE"].strip().replace("/", "-").replace("\\", "-")
            nombre_archivo = os.path.join(carpeta_individual, f"{nombre_estudiante}.xlsx")
            fila.to_frame().T.to_excel(nombre_archivo, index=False, sheet_name="Hoja1")

            plantilla_doc = word.Documents.Open(plantilla)
            plantilla_doc.MailMerge.OpenDataSource(
                Name=nombre_archivo,
                ConfirmConversions=False,
                ReadOnly=True,
                LinkToSource=True,
                AddToRecentFiles=False,
                Revert=False,
                Format=0,
                Connection='Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + nombre_archivo + ';Extended Properties="Excel 12.0 XML;HDR=YES";',
                SQLStatement='SELECT * FROM [Hoja1$]'
            )

            plantilla_doc.MailMerge.Destination = 0
            plantilla_doc.MailMerge.Execute(Pause=False)
            plantilla_doc.Close(False)

            combined_doc = word.ActiveDocument
            pdf_individual = os.path.join(carpeta_individual, f"{nombre_estudiante}.pdf")
            print(f"📄 Guardando PDF individual: {pdf_individual}")
            combined_doc.ExportAsFixedFormat(pdf_individual, 17)
            combined_doc.Close(False)

            

except KeyboardInterrupt as e:
    print(f"\n🛑 {e}")
except Exception as e:
    print(f"\n❌ Error inesperado: {e}")
finally:
    print("🧹 Cerrando Word...")
    word.Quit()
    tiempo_total = time.time() - start_time
    print(f"✅ Finalizado en {tiempo_total:.2f} segundos.")
