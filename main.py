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

# ------------------ UTILIDADES ------------------
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


def obtener_base(nombre_archivo):
    return nombre_archivo.split("_")[0].lower().strip()

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
        fila_actual = ws.max_row + 3

    nombre_salida = f"{base}_ReporteUnificado.xlsx"
    wb.save(os.path.join(ruta_salida, nombre_salida))
    print(f"✅ Archivo generado: {nombre_salida}")
