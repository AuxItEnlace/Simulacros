import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import statistics

# ---------------------- CONFIGURACION ----------------------

def leer_configuracion(filepath='config.ini'):
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            lines = file.read().splitlines()
        return {k.strip(): v.strip() for line in lines if '=' in line for k, v in [line.split('=', 1)]}
    except FileNotFoundError:
        raise FileNotFoundError("El archivo config.ini no fue encontrado.")

# ---------------------- PROCESAMIENTO DE DATOS ----------------------

def cargar_y_ordenar_datos(filepath):
    data = pd.read_csv(filepath, encoding="utf-8", engine="python")
    if data.empty:
        raise ValueError("El archivo CSV está vacío.")

    data[['GRADO_NUM', 'GRADO_LETRA']] = data['GRADO'].str.extract(r'(\d+)([A-Z])')
    data['GRADO_NUM'] = data['GRADO_NUM'].astype(int)
    data['% CORRECTAS'] = data['% CORRECTAS'].astype(str).str.replace('%', '').astype(float)

    data.sort_values(by=['GRADO_NUM', 'GRADO_LETRA', '% CORRECTAS', 'NOMBRE'],
                     ascending=[True, True, False, True], inplace=True)

    data.drop(columns=['GRADO_NUM', 'GRADO_LETRA'], inplace=True)
    
    return data

def calcular_columnas(data):
    # Calcular puntaje ponderado
    data['PUNTAJE_PONDERADO'] = (
        data['Correctas C1'] * 1 +
        data['Correctas C2'] * 2 +
        data['Correctas C3'] * 3 +
        data['CORRECTAS'] * 4
    )

    
    
    
    data.insert(0, 'PUESTO', data.groupby('GRADO')['PUNTAJE_PONDERADO']
                 .rank(ascending=False).astype(int))  # método 'first' evita empates
    
    

    # Eliminar columna auxiliar si no quieres que quede
    data.drop(columns=['PUNTAJE_PONDERADO'], inplace=True)

    # Extraer GRADO_NUM y GRADO_LETRA nuevamente para ordenar
    data[['GRADO_NUM', 'GRADO_LETRA']] = data['GRADO'].str.extract(r'(\d+)([A-Z])')
    data['GRADO_NUM'] = data['GRADO_NUM'].astype(int)

    # Ordenar por grado y luego por puesto (de menor a mayor)
    data.sort_values(by=['GRADO_NUM', 'GRADO_LETRA', 'PUESTO'],
                     ascending=[True, True, True], inplace=True)

    # Eliminar columnas auxiliares
    data.drop(columns=['GRADO_NUM', 'GRADO_LETRA'], inplace=True)

    # Insertar columnas respetando el orden
    data.insert(0, 'CONTAR', data.groupby('GRADO').cumcount(ascending=False) + 1)
    data.insert(1, 'PERCENTIL', data.groupby('GRADO')['% CORRECTAS'].transform(lambda x: (x.rank(pct=True) * 100).astype(int)))
    
    data.insert(2, 'PUNTAJE', data['% CORRECTAS'].round(2))

    # Mostrar % CORRECTAS como string con '%'
    data['% CORRECTAS'] = (data['% CORRECTAS'].astype(int).astype(str) + '%')

    return data



# ---------------------- CREACION DE HOJAS ----------------------

def crear_hojas_base(wb, nombres):
    ws = wb.active
    ws.title = nombres[0]
    for name in nombres[1:]:
        wb.create_sheet(name)
    return wb

# ---------------------- HOJAS DE GRADO ----------------------

def procesar_hojas_por_grado(wb, data, Sheets):
    avg = []
    data['GRADO_NUM'] = data['GRADO'].str.extract(r'^(\d+)').astype(int)

    for sheet_name in Sheets[3:]:
        grado_num = int(sheet_name.split('(')[1].split(')')[0])
        grado_data = data[data['GRADO_NUM'] == grado_num].copy()
        grado_data.drop(columns=['GRADO_NUM'], inplace=True)

        columns_to_drop = [str(i) for i in range(1, 43)] + ['DOCUMENTO'] + [f'Explicación {i}' for i in range(1, 43)] + [
            'INCORRECTAS', '% INCORRECTAS', 'INVÁLIDAS', '% INVÁLIDAS', 'VACÍAS', '% VACÍAS']
        grado_data.drop(columns=columns_to_drop, inplace=True, errors='ignore')

        int_avg_cols = ['CORRECTAS'] + [f'Correctas C{i}' for i in range(1, 4)] + [f'Correctas EI{i}' for i in range(1, 13)]
        porc_avg_cols = ['% CORRECTAS'] + [f'% Correctas C{i}' for i in range(1, 4)] + [f'% Correctas EI{i}' for i in range(1, 13)]
        avg_cols = [x for pair in zip(int_avg_cols, porc_avg_cols) for x in pair]

        sheet = wb[sheet_name]
        sheet.append(grado_data.columns.tolist())

        for grado, group in grado_data.groupby('GRADO'):
            for r in group.itertuples(index=False):
                sheet.append(list(r))

            group_valid = group[group['CORRECTAS'] > 0]
            avg_row = []
            for col in grado_data.columns:
                if col in avg_cols:
                    if col in int_avg_cols:
                        avg_row.append(round(group_valid[col].astype(float).mean(), 2))
                    else:
                        avg_row.append(f'{round(group_valid[col].astype(str).str.replace("%", "").astype(float).mean(), 2)}%')
                elif col == 'GRADO':
                    avg_row.append(f'PROMEDIO {grado}')
                else:
                    avg_row.append("")

            header_avg = ["", "", "", "", "", "", "TOTAL ESTUDIANTES", len(group_valid), "DESVIACION ESTANDAR",
                          round(group_valid['% CORRECTAS'].astype(str).str.replace('%', '').astype(float).std(), 2)]
            sheet.append(header_avg)
            sheet.append(avg_row)

            resumen = avg_row[6:]
            resumen[0] = grado
            #Desviación estándar
            resumen[1] = round(group_valid['% CORRECTAS'].astype(str).str.replace('%', '').astype(float).std(), 2)
            avg.append(resumen)

    return avg

# ---------------------- FORMATEO FINAL ----------------------

def formatear_porcentajes(wb, Sheets):
    for sheet_name in Sheets:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=sheet.max_column):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip().endswith('%'):
                    try:
                        cell.value = float(cell.value.strip().replace('%', '')) / 100
                        cell.number_format = '0%'
                    except ValueError:
                        pass
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = '0' if cell.value == int(cell.value) else '0.0'

# ---------------------- FUNCION PARA CALCULAR PERCENTILES SIN scipy ----------------------

def calcular_percentiles(puntajes):
    n = len(puntajes)
    sorted_puntajes = sorted(puntajes)
    
    percentiles = []
    for p in puntajes:
        posiciones = [i + 1 for i, val in enumerate(sorted_puntajes) if val == p]
        rango_promedio = sum(posiciones) / len(posiciones)
        percentil = round((rango_promedio / n) * 100)
        percentiles.append(percentil)
    return percentiles

# ---------------------- FUNCION PARA AGREGAR RESUMEN POR GRUPO ----------------------

def agregar_resumen_por_grupo(sheet, avg_sorted, colegio):
    grupos = {}
    for row in avg_sorted:
        grado = row[0]
        numero = int(re.match(r"(\d+)", grado).group(1))
        grupos.setdefault(numero, []).append(list(row))

    promedios_grupos = []
    for numero_grupo, filas in grupos.items():
        # Extraer puntajes para percentil y puesto (asumo la posición 2 para porcentaje)
        puntajes = []
        for fila in filas:
            valor = str(fila[2]).replace('%', '').replace(',', '.')
            try:
                puntajes.append(float(valor))
            except:
                puntajes.append(0.0)

        percentiles = calcular_percentiles(puntajes)
        puestos = [sorted(puntajes, reverse=True).index(p) + 1 for p in puntajes]

        # Agregar columnas Percentil y Puesto al inicio de cada fila
        filas_modificadas = []
        for i, fila in enumerate(filas):
            # Calculamos promedio desde la columna 3 en adelante (índice 3)
            valores_para_promedio = []
            for val in fila[1:]:
                try:
                    v = float(str(val).replace('%','').replace(',','.'))
                    valores_para_promedio.append(v)
                except:
                    pass
            promedio_fila = round(sum(valores_para_promedio)/len(valores_para_promedio), 2) if valores_para_promedio else 0
            
            fila_nueva = [percentiles[i], puestos[i]] + fila
            filas_modificadas.append(fila_nueva)
            filas_modificadas = sorted(filas_modificadas, key=lambda x: x[1])

        # Escribir filas con percentil, puesto y promedio fila al final
        for fila in filas_modificadas:
            sheet.append(fila)

        # Posición donde está el total de correctas (después de columna vacía)
        # Buscamos columna vacía y tomamos la siguiente columna para desviación
        primera_fila = filas_modificadas[0]
        try:
            idx_vacio = primera_fila.index('')
            idx_total_correctas = idx_vacio + 1
        except ValueError:
            # Si no hay columna vacía, asumimos índice 3 para total correctas
            idx_total_correctas = 2

        # Extraemos valores para desviación solo de la columna total correctas
        valores_correctas = []
        for fila in filas_modificadas:
            try:
                val = fila[idx_total_correctas+2]
                if isinstance(val, str):
                    val = val.replace('%','').replace(',','.')
                valores_correctas.append(float(val))
            except:
                valores_correctas.append(0.0)

        desviacion_total_correctas = round(statistics.stdev(valores_correctas), 2) if len(valores_correctas) > 1 else 0

        # Construir fila desviación estándar:
        # Poner vacío en Percentil y Puesto, texto en columna 3,
        # vacíos hasta la columna total correctas, desviación allí, vacíos para resto de columnas, vacíos para promedio fila
        fila_desviacion = [''] * (idx_total_correctas + 3)  # CORRECCIÓN: aumentar tamaño para evitar IndexError
        fila_desviacion[2] = 'Desviación Estándar'  # Columna 3 (índice 2)
        fila_desviacion[idx_total_correctas+1] = desviacion_total_correctas  # Ajustamos índice sumando 2 por columnas extra al inicio
        sheet.append(fila_desviacion)

        # Construir fila promedio para cada columna desde idx_total_correctas en adelante
        promedios_cols = []
        num_cols = len(filas_modificadas[0])

        for col_idx in range(idx_total_correctas, num_cols):  # columnas desde total correctas en adelante, incluyendo promedio fila al final
            valores_col = []
            for fila in filas_modificadas:
                try:
                    val = fila[col_idx]
                    if isinstance(val, str):
                        val = val.replace('%','').replace(',','.')
                    valores_col.append(float(val))
                except:
                    pass
            promedio_col = round(sum(valores_col)/len(valores_col), 2) if valores_col else 0
            # Si la columna original era porcentaje, formateamos con %
            if (isinstance(filas_modificadas[0][col_idx], str) and filas_modificadas[0][col_idx].endswith('%')):
                promedio_col = f"{promedio_col}%"
            promedios_cols.append(promedio_col)

        # Construir fila promedio
        fila_promedio = [''] * (idx_total_correctas)

        fila_promedio.extend(promedios_cols)
        fila_promedio[2] = 'Promedio'
        sheet.append(fila_promedio)
        promedios_grupos.append([colegio] + [f"{numero_grupo}°"] +[desviacion_total_correctas] + promedios_cols[2:])
    return promedios_grupos

# ---------------------- FUNCION PRINCIPAL ----------------------

def main():
    config = leer_configuracion()
    data = cargar_y_ordenar_datos('data/data.csv')
    data = calcular_columnas(data)

    school_name = config["Sede"][:10] if len(config["Sede"]) > 10 else config["Sede"]
    Sheets = ['Reprorte General', 'Detallado Grados', f'BD-RESULTADOS-{school_name}'] + [f'GRADO ({i})' for i in range(12)]
    wb = crear_hojas_base(Workbook(), Sheets)

    # Cargar hoja de resultados
    for r in dataframe_to_rows(data, index=False, header=True):
        wb[Sheets[2]].append(r)

    encabezado = ['COLEGIO','GRADO', 'DESVESTA', 'CORRECTAS', '%CORRECTAS']

    # Intercalar Correctas C1, % Correctas C1, ..., C3
    for i in range(1, 4):
        encabezado.append(f'Correctas C{i}')
        encabezado.append(f'% Correctas C{i}')

    # Intercalar Correctas EI1, % Correctas EI1, ..., EI12
    for i in range(1, 13):
        encabezado.append(f'Correctas EI{i}')
        encabezado.append(f'% Correctas EI{i}')

    wb[Sheets[0]].append(encabezado)
    encabezado = ['PERCENTIL', 'PUESTO', 'GRADO', 'DESVESTA', 'CORRECTAS', '%CORRECTAS']

    # Intercalar Correctas C1, % Correctas C1, ..., C3
    for i in range(1, 4):
        encabezado.append(f'Correctas C{i}')
        encabezado.append(f'% Correctas C{i}')

    # Intercalar Correctas EI1, % Correctas EI1, ..., EI12
    for i in range(1, 13):
        encabezado.append(f'Correctas EI{i}')
        encabezado.append(f'% Correctas EI{i}')

    wb[Sheets[1]].append(encabezado)

    avg = procesar_hojas_por_grado(wb, data, Sheets)
    avg_sorted = sorted(avg, key=lambda x: int(re.match(r"(\d+)", x[0]).group(1)))

    avg_grados = agregar_resumen_por_grupo(wb[Sheets[1]], avg_sorted, config['SchoolName'] + ' ' + config['Sede'])
    #calcular desviacion estandar de los promedios de los grados columna 4
    valores = [fila[3] for fila in avg_grados]
    print(f"Valores para desviación estándar: {valores}")
    if len(valores) > 1:
        desviacion_total_correctas = round(statistics.stdev(valores), 2)
    else:
        desviacion_total_correctas = 0
    for row in avg_grados:
        wb[Sheets[0]].append(row)
    
    wb[Sheets[0]].append(['Desvacion Estandar del Colegio', desviacion_total_correctas])
    
    
    #formatear_porcentajes(wb, Sheets)

    wb.save(f'data/{config["SchoolName"]} {config["Sede"]} {config["Date"]} PRUEBA {config["Prueba"]}.xlsx')

if __name__ == "__main__":
    main()
        