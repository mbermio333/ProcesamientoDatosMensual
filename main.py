# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage

# ------------------ CONFIGURACIÓN GENERAL ------------------
ruta_fm = "MedicionesFmCSV"
ruta_tv = "MedicionesTvCSV"
ruta_salida = "ReportesUnificados"
ruta_imagenes = "Img"
os.makedirs(ruta_salida, exist_ok=True)
fecha_actual = datetime.today().strftime("%d/%m/%Y")

# ------------------ FUNCIONES AUXILIARES ------------------

def combinar_observaciones_solo_en_tv(ws, fila_inicio_tabla, fila_fin_tabla, num_columnas, es_tv=True):
    """
    Combina horizontalmente la última y penúltima columna para cada fila SOLO si es TV.
    """
    if es_tv:
        for fila in range(fila_inicio_tabla + 1, fila_fin_tabla + 1):
            ws.merge_cells(start_row=fila, start_column=num_columnas - 1, end_row=fila, end_column=num_columnas)
        
        # También combinar el encabezado (la celda con el nombre del campo)
        ws.merge_cells(start_row=fila_inicio_tabla, start_column=num_columnas - 1, end_row=fila_inicio_tabla, end_column=num_columnas)
        celda = ws.cell(row=fila_inicio_tabla, column=num_columnas - 1)
        celda.value = "OBSERVACIONES"
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(row=fila_inicio_tabla, column=num_columnas).value = None

def formatear_hoja(ws, fila_inicio, encabezado_lineas):
    num_columnas = ws.max_column
    ws.insert_rows(fila_inicio, amount=len(encabezado_lineas))

    for i, texto in enumerate(encabezado_lineas):
        fila = fila_inicio + i
        celda = ws.cell(row=fila, column=1, value=texto)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=num_columnas)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.font = Font(bold=True, size=12)

    # Insertar imágenes
    logo_izquierda = os.path.join(ruta_imagenes, "ARCOTEL.png")
    logo_derecha = os.path.join(ruta_imagenes, "nEcuador.png")

    fila_img_izq = fila_inicio + 2
    fila_img_der = fila_inicio + 2

    if os.path.exists(logo_izquierda):
        img_left = XLImage(logo_izquierda)
        img_left.width = 500
        img_left.height = 100
        ws.add_image(img_left, f"A{fila_img_izq}")

    if os.path.exists(logo_derecha):
        img_right = XLImage(logo_derecha)
        img_right.width = 260
        img_right.height = 120
        ws.add_image(img_right, f"AH{fila_img_der}")

    # Bordes
    inicio_fila_tabla = fila_inicio + len(encabezado_lineas)
    fin_fila_tabla = ws.max_row
    thin = Side(border_style="thin")
    borde_grueso = Side(border_style="medium")

    for row in range(inicio_fila_tabla, fin_fila_tabla + 1):
        for col in range(1, num_columnas + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            if row == inicio_fila_tabla:
                cell.border = Border(top=borde_grueso, bottom=thin,
                                     left=borde_grueso if col == 1 else thin,
                                     right=borde_grueso if col == num_columnas else thin)
            elif row == fin_fila_tabla:
                cell.border = Border(top=thin, bottom=borde_grueso,
                                     left=borde_grueso if col == 1 else thin,
                                     right=borde_grueso if col == num_columnas else thin)
            elif col == 1:
                cell.border = Border(left=borde_grueso, top=thin, bottom=thin, right=thin)
            elif col == num_columnas:
                cell.border = Border(right=borde_grueso, top=thin, bottom=thin, left=thin)

            cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, num_columnas + 1):
        celda = ws.cell(row=inicio_fila_tabla, column=col)
        if celda.value == "Promedio(dBuV/m)":
            celda.value = "Promedio\n(dBuV/m)"
            ws.column_dimensions[get_column_letter(col)].width = 15
        elif celda.value == "Ancho de Banda (KHz)":
            celda.value = "Ancho de Banda\n(KHz)"
        elif celda.value == "Medición Manual":
            celda.value = "Medición Manual\nAB(KHz) o NIVEL\n(dBµV/m)"
            ws.column_dimensions[get_column_letter(col)].width = 23
        celda.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    # Ajustar anchos del resto
    for i in range(1, num_columnas + 1):
        if ws.column_dimensions[get_column_letter(i)].width in [15, 23]:
            continue
        col_letter = get_column_letter(i)
        max_length = 0
        for cell in ws[col_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    ws.row_dimensions[inicio_fila_tabla].height = 45
# ------------------ FUNCIÓN PARA COLOREAR CELDAS ------------------
def colorear_celdas_por_valor(ws, fila_inicio, fila_fin, columna_inicio, columna_fin):
    """
    Colorea las celdas según su valor numérico:
    - Verde: 0 a 100
    - Amarillo: 100 a 200  
    - Rojo: mayor a 200
    """
    from openpyxl.styles import PatternFill
    
    # Colores corregidos con formato ARGB (FF + código RGB)
    verde = PatternFill(start_color="FFCDFECE", end_color="FFCDFECE", fill_type="solid")
    amarillo = PatternFill(start_color="FFFFFE9F", end_color="FFFFFE9F", fill_type="solid")
    rosa = PatternFill(start_color="FFFD9BCB", end_color="FFFD9BCB", fill_type="solid")
    rojo = PatternFill(start_color="FFFC4A2C", end_color="FFFC4A2C", fill_type="solid")  # por si lo necesitas también

    # Iterar por todas las celdas en el rango
    for fila in range(fila_inicio, fila_fin + 1):
        for col in range(columna_inicio, columna_fin + 1):
            celda = ws.cell(row=fila, column=col)
            
            # Verificar si la celda tiene un valor numérico
            try:
                if celda.value is not None and str(celda.value).replace('.', '', 1).isdigit():
                    valor = float(celda.value)
                    
                    # Aplicar color según el rango
                    if 0 <= valor <= 43:
                        celda.fill = rosa
                    elif 43 < valor < 54:
                        celda.fill = amarillo
                    elif valor >= 54:
                        celda.fill = verde
                        
            except (ValueError, TypeError):
                # Si no es número, no hacer nada
                pass

# ------------------ FUNCIÓN PARA ENCONTRAR COLUMNAS NUMÉRICAS ------------------
def encontrar_columnas_numericas(ws, fila_encabezados):
    """
    Encuentra automáticamente las columnas que contienen valores numéricos
    basándose en los encabezados de días (1 al 31)
    """
    columnas_numericas = []
    
    for col in range(1, ws.max_column + 1):
        celda = ws.cell(row=fila_encabezados, column=col)
        if celda.value and str(celda.value).isdigit():
            try:
                dia = int(celda.value)
                if 1 <= dia <= 31:
                    columnas_numericas.append(col)
            except ValueError:
                pass
    
    if columnas_numericas:
        return min(columnas_numericas), max(columnas_numericas)
    else:
        # Valores por defecto si no encuentra días
        return 3, 33

    # --- 🔁 Combinación solo para la tabla TV ---
    nombre_hoja = ws.cell(row=fila_inicio + len(encabezado_lineas) - 1, column=1).value
    if "TV" in nombre_hoja:
        col_observaciones = num_columnas
        col_manual = num_columnas - 1
        ancho_observaciones = ws.column_dimensions[get_column_letter(col_observaciones)].width
        ancho_manual = ws.column_dimensions[get_column_letter(col_manual)].width

        for fila in range(inicio_fila_tabla + 1, fin_fila_tabla + 1):
            ws.merge_cells(start_row=fila, start_column=col_manual, end_row=fila, end_column=col_observaciones)
        ws.merge_cells(start_row=inicio_fila_tabla, start_column=col_manual, end_row=inicio_fila_tabla, end_column=col_observaciones)

        # Reemplazar el texto del encabezado
        celda = ws.cell(row=inicio_fila_tabla, column=col_manual)
        celda.value = "OBSERVACIONES"
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(row=inicio_fila_tabla, column=col_observaciones).value = None

        # Asignar el nuevo ancho sumado
        ws.column_dimensions[get_column_letter(col_manual)].width = ancho_observaciones + ancho_manual


def obtener_base(nombre_archivo):
    return nombre_archivo.split("_")[0].lower().strip()

# ------------------ PROCESAMIENTO ------------------

archivos_fm = {obtener_base(f): os.path.join(ruta_fm, f) for f in os.listdir(ruta_fm) if f.endswith(".csv")}
archivos_tv = {obtener_base(f): os.path.join(ruta_tv, f) for f in os.listdir(ruta_tv) if f.endswith(".csv")}

nombres_bases = set(archivos_fm.keys()).union(archivos_tv.keys())

for base in nombres_bases:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    fila_actual = 1

    for tipo, archivos, titulo in [
        ("FM", archivos_fm, "FORMULARIO DE CONTROL MENSUAL DE FM"),
        ("TV", archivos_tv, "FORMULARIO DE CONTROL MENSUAL DE TV")
    ]:
        if base not in archivos:
            continue

        ruta_archivo = archivos[base]
        df = pd.read_csv(ruta_archivo, encoding="unicode_escape")
        df = df.rename(columns={"Nombre de la estación": "ESTACION"}) if tipo == "TV" else df
        df = df[df.columns[:9]] if tipo == "FM" else df[df.columns[:5]]

        df["ESTACION"] = df["ESTACION"].astype(str).str.strip()
        df = df[df["ESTACION"].notna() & (df["ESTACION"] != "")]
        df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date
        mes_objetivo = df["Tiempo"].dropna().apply(lambda x: x.month).value_counts().idxmax()

        nombre_mes_es = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"][mes_objetivo - 1]
        df = df[df["Tiempo"].apply(lambda x: x.month == mes_objetivo)]
        df["DIA"] = pd.to_datetime(df["Tiempo"]).dt.day

        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"] / 1_000_000
        df["Level (dBµV/m)"] = df["Level (dBµV/m)"].astype(str).str.replace(",", ".", regex=False).astype(float)

        agrupado = df.groupby(["ESTACION", "Frecuencia (MHz)", "DIA"])["Level (dBµV/m)"].mean().reset_index()
        pivot = agrupado.pivot(index=["ESTACION", "Frecuencia (MHz)"], columns="DIA", values="Level (dBµV/m)")
        todos_los_dias = list(range(1, 32))
        pivot = pivot.reindex(columns=todos_los_dias, fill_value=0)

        pivot[todos_los_dias] = pivot[todos_los_dias].replace(0, "-")
        pivot_numeric = pivot[todos_los_dias].replace("-", pd.NA).apply(pd.to_numeric, errors="coerce")
        pivot["Promedio(dBuV/m)"] = pivot_numeric.mean(axis=1, skipna=True).round(2)

        if tipo == "FM":
            ancho_banda = (
                df.groupby(["ESTACION", "Frecuencia (MHz)"])["Bandwidth (Hz)"].mean().reset_index()
            )
            ancho_banda["Ancho de Banda (KHz)"] = (ancho_banda["Bandwidth (Hz)"] / 1000).round(2)
            pivot = pivot.reset_index().merge(ancho_banda.drop(columns=["Bandwidth (Hz)"]), on=["ESTACION", "Frecuencia (MHz)"], how="left")
        else:
            pivot = pivot.reset_index()

        pivot["Medición Manual"] = ""
        pivot["OBSERVACIONES"] = ""
        pivot = pivot.sort_values(by="Frecuencia (MHz)")

        for col in pivot.select_dtypes(include="number").columns:
            pivot[col] = pivot[col].round(2)

        for i, row in enumerate(dataframe_to_rows(pivot, index=False, header=True)):
            for j, val in enumerate(row, start=1):
                ws.cell(row=fila_actual + i, column=j, value=val)

        encabezado_fm = [
            "INFORME DE CONTROL TÉCNICO",
            "No. IT-CZ06-R-2025-00XX",
            "AGENCIA DE REGULACIÓN Y CONTROL DE LAS TELECOMUNICACIONES",
            "COORDINACIÓN ZONAL 6",
            "ESTACIÓN DE COMPROBACIÓN TÉCNICA",
            titulo,
            "CIUDAD:" + ("tambo" if base == "cañar" else base.upper()),
            f"PERIODO: {nombre_mes_es.upper()}",
            f"FECHA PRESENTACIÓN: {fecha_actual}"
        ] if tipo == "FM" else [
            "AGENCIA DE REGULACIÓN Y CONTROL DE LAS TELECOMUNICACIONES",
            "COORDINACIÓN ZONAL 6",
            "ESTACIÓN DE COMPROBACIÓN TÉCNICA",
            titulo,
            "CIUDAD:" + ("tambo" if base == "cañar" else base.upper()),
            f"PERIODO: {nombre_mes_es.upper()}",
            f"FECHA PRESENTACIÓN: {fecha_actual}"
        ]

        formatear_hoja(ws, fila_actual, encabezado_fm)

        # ✅ ENCONTRAR Y COLOREAR CELDAS NUMÉRICAS
        # Calcular fila de encabezados de columnas (donde dice "ESTACION", "Frecuencia", etc.)
        fila_encabezados_columnas = fila_actual + len(encabezado_fm)

        # Encontrar columnas con valores numéricos (días 1-31)
        col_inicio_numeros, col_fin_numeros = encontrar_columnas_numericas(ws, fila_encabezados_columnas)

        # Calcular filas de datos (después de los encabezados de columnas)
        fila_inicio_datos = fila_encabezados_columnas + 1
        fila_fin_datos = fila_inicio_datos + len(pivot) - 1

        # Aplicar colores a las celdas numéricas
        colorear_celdas_por_valor(ws, fila_inicio_datos, fila_fin_datos, 
                                col_inicio_numeros, col_fin_numeros)

        fila_actual = ws.max_row + 3

    nombre_salida = f"{base}_ReporteUnificado.xlsx"
       
    from openpyxl.utils import get_column_letter

    for sheet in wb.worksheets:
        if "Resumen" not in sheet.title:
            continue

        # Buscar fila del segundo encabezado (TV)
        for fila in range(1, sheet.max_row + 1):
            val = sheet.cell(row=fila, column=1).value
            if isinstance(val, str) and "FORMULARIO DE CONTROL MENSUAL DE TV" in val:
                fila_encabezado_tv = fila
                break
        else:
            continue  # No hay TV en esta hoja

        # Buscar fila exacta de la cabecera de la tabla (donde dice "ESTACION")
        fila_tabla_tv = None
        for fila in range(fila_encabezado_tv + 1, sheet.max_row + 1):
            if sheet.cell(row=fila, column=1).value == "ESTACION":
                fila_tabla_tv = fila
                break

        if fila_tabla_tv is None:
            continue  # Seguridad por si no se encuentra

        ultima_fila_tv = sheet.max_row
        col_final = sheet.max_column
        col_penultima = col_final - 1

        # Obtener anchos actuales
        ancho_col1 = sheet.column_dimensions[get_column_letter(col_penultima)].width
        ancho_col2 = sheet.column_dimensions[get_column_letter(col_final)].width
        ancho_combinado = (ancho_col1 or 10) + (ancho_col2 or 10)

        # Combinar celdas del cuerpo de la tabla
        for fila in range(fila_tabla_tv + 1, ultima_fila_tv + 1):
            sheet.merge_cells(start_row=fila, start_column=col_penultima, end_row=fila, end_column=col_final)

        # Combinar encabezado de tabla ("OBSERVACIONES")
        sheet.merge_cells(start_row=fila_tabla_tv, start_column=col_penultima, end_row=fila_tabla_tv, end_column=col_final)
        celda_obs = sheet.cell(row=fila_tabla_tv, column=col_penultima)
        celda_obs.value = "OBSERVACIONES"
        celda_obs.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                # Combinar encabezado de tabla ("OBSERVACIONES")
        sheet.merge_cells(start_row=fila_tabla_tv, start_column=col_penultima, end_row=fila_tabla_tv, end_column=col_final)

        # Solo escribir en la celda principal de la combinación
        celda_obs = sheet.cell(row=fila_tabla_tv, column=col_penultima)
        celda_obs.value = "OBSERVACIONES"
        celda_obs.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# No escribas en col_final (ya está combinada)


        # Asignar nuevo ancho
        sheet.column_dimensions[get_column_letter(col_penultima)].width = ancho_combinado

    
    wb.save(os.path.join(ruta_salida, nombre_salida))
    print(f"✅ Archivo generado: {nombre_salida}")
