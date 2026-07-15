# -*- coding: utf-8 -*-
import os
import glob
import pandas as pd

def merge_txt_to_csv():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "input")
    output_dir = os.path.join(base_dir, "output")

    print(f"Buscando archivos .txt y .csv en: {input_dir}")

    if not os.path.isdir(input_dir):
        print("No existe la carpeta 'input'. Asegúrate de crearla y colocar los archivos .txt o .csv allí.")
        return

    os.makedirs(output_dir, exist_ok=True)

    archivos_txt = glob.glob(os.path.join(input_dir, "*.txt"))
    archivos_csv = glob.glob(os.path.join(input_dir, "*.csv"))
    archivos_entrada = sorted(archivos_txt + archivos_csv)

    if not archivos_entrada:
        print("No se encontraron archivos .txt ni .csv en la carpeta 'input'.")
        return

    archivo_salida = "resultado_combinado.csv"
    lista_dataframes = []

    for ruta_archivo in archivos_entrada:
        nombre_archivo = os.path.basename(ruta_archivo)
        nombre_sin_ext, extension = os.path.splitext(nombre_archivo)

        print(f"Procesando: {nombre_archivo}...")

        try:
            if extension.lower() == ".txt":
                df = pd.read_csv(ruta_archivo, sep=';', dtype=str, keep_default_na=False)
            else:
                try:
                    df = pd.read_csv(ruta_archivo, sep=None, engine='python', dtype=str, keep_default_na=False)
                except Exception:
                    df = pd.read_csv(ruta_archivo, sep=';', dtype=str, keep_default_na=False)

            df.insert(0, 'Origen', nombre_sin_ext)
            lista_dataframes.append(df)

        except Exception as e:
            print(f"Error al procesar el archivo {nombre_archivo}: {e}")

    if not lista_dataframes:
        print("No se pudo procesar ningún archivo válido.")
        return

    print("Combinando datos y alineando columnas...")
    df_final = pd.concat(lista_dataframes, axis=0, ignore_index=True, sort=False)

    ruta_salida = os.path.join(output_dir, archivo_salida)
    df_final.to_csv(ruta_salida, sep=';', index=False, encoding='utf-8-sig')

    print("¡Proceso completado con éxito!")
    print(f"Archivo guardado en: {ruta_salida}")

if __name__ == "__main__":
    merge_txt_to_csv()