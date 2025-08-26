import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font

# Directorios de entrada y salida
directorio_fm = "Promedios_mensuales"
directorio_tv = "pruebas"
directorio_salida = "ReportesUnificados"
os.makedirs(directorio_salida, exist_ok=True)

# Obtener archivos Excel
archivos_fm = {f.split("_")[0]: os.path.join(directorio_fm, f) for f in os.listdir(directorio_fm) if f.endswith(".xlsx")}
archivos_tv = {f.split("_")[0]: os.path.join(directorio_tv, f) for f in os.listdir(directorio_tv) if f.endswith(".xlsx")}

# Iterar sobre los nombres coincidentes
for nombre in archivos_fm.keys() & archivos_tv.keys():
    ruta_fm = archivos_fm[nombre]
    ruta_tv = archivos_tv[nombre]

    # Leer hojas de FM y TV como DataFrames
    df_fm = pd.read_excel(ruta_fm, sheet_name="Resumen")
    df_tv = pd.read_excel(ruta_tv, sheet_name="Resumen")

    # Crear archivo combinado temporal
    archivo_salida = os.path.join(directorio_salida, f"{nombre}_reporte_unificado.xlsx")

    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        # Escribir FM
        df_fm.to_excel(writer, index=False, sheet_name="Resumen", startrow=0)

        # Escribir TV debajo de FM, con dos filas de separación
        start_row = len(df_fm) + 3
        df_tv.to_excel(writer, index=False, sheet_name="Resumen", startrow=start_row)

    # Aplicar formato opcional con openpyxl
    wb = load_workbook(archivo_salida)
    ws = wb["Resumen"]

    # Centrar todos los valores
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Ajustar ancho de columnas automáticamente
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_length = max((len(str(cell.value)) if cell.value else 0) for cell in ws[col_letter])
        ws.column_dimensions[col_letter].width = max_length + 2

    # Título
    ws.insert_rows(1)
    ws.cell(row=1, column=1, value=f"REPORTE UNIFICADO DE FM Y TV - {nombre.upper()}").font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    # Guardar
    wb.save(archivo_salida)
    print(f"✅ Archivo combinado guardado: {archivo_salida}")
