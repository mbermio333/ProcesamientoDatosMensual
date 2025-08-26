
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# --- CONFIGURACIONES GENERALES ---
ruta_fm = "Promedios_mensuales"
ruta_tv = "pruebas"
ruta_salida = "ReportesUnificados"
ruta_imagenes = "Img"
os.makedirs(ruta_salida, exist_ok=True)
fecha_actual = datetime.today().strftime("%d/%m/%Y")

# --- UTILITARIOS ---
def formatear_columna_excel(ws):
    num_columnas = ws.max_column
    for i in range(1, num_columnas + 1):
        col_letter = get_column_letter(i)
        max_length = 0
        for cell in ws[col_letter]:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

def aplicar_bordes(ws, inicio_fila, fin_fila):
    num_columnas = ws.max_column
    thin = Side(border_style="thin")
    medium = Side(border_style="medium")
    for row in range(inicio_fila, fin_fila + 1):
        for col in range(1, num_columnas + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            if row == inicio_fila:
                cell.border = Border(top=medium, bottom=cell.border.bottom, left=cell.border.left, right=cell.border.right)
            if row == fin_fila:
                cell.border = Border(top=cell.border.top, bottom=medium, left=cell.border.left, right=cell.border.right)
            if col == 1:
                cell.border = Border(top=cell.border.top, bottom=cell.border.bottom, left=medium, right=cell.border.right)
            if col == num_columnas:
                cell.border = Border(top=cell.border.top, bottom=cell.border.bottom, left=cell.border.left, right=medium)

def centrar_y_saltos_linea(ws, fila_encabezado):
    for row in ws.iter_rows(min_row=fila_encabezado, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila_encabezado].height = 45

def insertar_encabezado(ws, encabezado, fila_inicio):
    num_columnas = ws.max_column
    for i, texto in enumerate(encabezado):
        ws.insert_rows(fila_inicio + i)
        celda = ws.cell(row=fila_inicio + i, column=1, value=texto)
        ws.merge_cells(start_row=fila_inicio + i, start_column=1, end_row=fila_inicio + i, end_column=num_columnas)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.font = Font(bold=True, size=12)
    return fila_inicio + len(encabezado)

def insertar_imagen(ws, path, celda):
    if os.path.exists(path):
        img = XLImage(path)
        img.width = 500
        img.height = 100
        img.anchor = celda
        img.left = 0
        img.top = 0
        ws.add_image(img)

# --- CARGAR REPORTES FM Y TV ---
reportes_fm = {f.split("_")[0]: os.path.join(ruta_fm, f) for f in os.listdir(ruta_fm) if f.endswith(".xlsx")}
reportes_tv = {f.split("_")[0]: os.path.join(ruta_tv, f) for f in os.listdir(ruta_tv) if f.endswith(".xlsx")}
coincidentes = set(reportes_fm.keys()).intersection(reportes_tv.keys())

for base in coincidentes:
    print(f"📄 Unificando: {base}")
    wb_unificado = Workbook()
    ws = wb_unificado.active
    ws.title = "Resumen"
    ws.sheet_view.showGridLines = False
    fila_actual = 1

    for tipo, path_excel in [("FM", reportes_fm[base]), ("TV", reportes_tv[base])]:
        wb_origen = load_workbook(path_excel)
        ws_origen = wb_origen.active
        datos = list(ws_origen.values)
        encabezado = []
        for i in range(len(datos)):
            fila = datos[i]
            if (
                any("CIUDAD" in str(c) for c in fila)
                 and i + 2 < len(datos)
                and any("PERIODO" in str(c) for c in datos[i + 1])
                and any("FECHA PRESENTACIÓN" in str(c) for c in datos[i + 2])
            ):
                encabezado = datos[:i + 3]  # Incluye hasta "FECHA PRESENTACIÓN"
                datos_tabla = datos[i + 3:]
                break


        if tipo == "TV":
            ws.append([])
            ws.append([])
            fila_actual = ws.max_row + 1

        fila_inicio_encabezado = fila_actual
        for linea in encabezado:
            ws.append([linea[0]])
            ws.merge_cells(start_row=fila_actual, start_column=1, end_row=fila_actual, end_column=len(datos_tabla[0]))
            celda = ws.cell(row=fila_actual, column=1)
            celda.alignment = Alignment(horizontal="center", vertical="center")
            celda.font = Font(bold=True, size=12)
            fila_actual += 1

        fila_img = fila_inicio_encabezado
        insertar_imagen(ws, os.path.join(ruta_imagenes, "ARCOTEL.png"), f"A{fila_img}")
        insertar_imagen(ws, os.path.join(ruta_imagenes, "nEcuador.png"), f"AH{fila_img}")

        for row in datos_tabla:
            ws.append(row)

        fila_inicio = fila_actual
        fila_fin = ws.max_row
        aplicar_bordes(ws, fila_inicio, fila_fin)
        centrar_y_saltos_linea(ws, fila_inicio)
        formatear_columna_excel(ws)

        fila_actual = ws.max_row + 1

    ruta_salida_final = os.path.join(ruta_salida, f"{base}_reporte_unificado.xlsx")
    wb_unificado.save(ruta_salida_final)
    print(f"✅ Archivo creado: {ruta_salida_final}")
