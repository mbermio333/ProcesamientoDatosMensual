import pandas as pd
import calendar
import os
from datetime import datetime


# ---------- CONFIGURACIÓN GENERAL ----------
ruta_entrada = "MedicionesFmCSV"
ruta_salida = "Promedios_mensuales"
os.makedirs(ruta_salida, exist_ok=True)
fecha_actual = datetime.today().strftime("%d/%m/%Y")

# ---------- PROCESAMIENTO DE CADA ARCHIVO ----------
for archivo_entrada in os.listdir(ruta_entrada):
    if archivo_entrada.endswith(".csv"):
        try:
            print(f"📄 Procesando archivo: {archivo_entrada}")

            ruta_csv_entrada = os.path.join(ruta_entrada, archivo_entrada)

            # Cargar datos
            df = pd.read_csv(ruta_csv_entrada, encoding='unicode_escape')
            df = df.iloc[:, :9]
            df["ESTACION"] = df["ESTACION"].astype(str).str.strip()
            df = df[df["ESTACION"].notna() & (df["ESTACION"] != "")]

            # Formatear columnas
            df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date

            # Detectar mes con más datos
            mes_ocurrencias = df["Tiempo"].apply(lambda x: x.month)
            mes_objetivo = mes_ocurrencias.value_counts().idxmax()

            nombre_mes_es = {
                1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
                5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
                9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
            }[mes_objetivo]

            # Filtrar por mes
            df["FRECUENCIA (MHz)"] = df["Frecuencia (Hz)"] / 1_000_000
            df["Level (dBµV/m)"] = df["Level (dBµV/m)"].astype(str).str.replace(",", ".", regex=False).astype(float)
            df = df[df["Tiempo"].apply(lambda x: x.month == mes_objetivo)]

            if df.empty:
                print(f"⚠ No hay datos válidos para el mes más frecuente ({nombre_mes_es}) en {archivo_entrada}")
            else:
                df["DIA"] = pd.to_datetime(df["Tiempo"]).dt.day

                agrupado = df.groupby(["ESTACION", "FRECUENCIA (MHz)", "DIA"])["Level (dBµV/m)"].mean().reset_index()
                pivot = agrupado.pivot(index=["ESTACION", "FRECUENCIA (MHz)"], columns="DIA", values="Level (dBµV/m)")

                # Rellenar días del 1 al 31
                todos_los_dias = list(range(1, 32))
                pivot = pivot.reindex(columns=todos_los_dias, fill_value=0)

                # Promedio mensual Level
                promedios_sin_ceros = pivot.replace(0, pd.NA).mean(axis=1, skipna=True)
                pivot["Promedio(dBuV/m)"] = promedios_sin_ceros.astype(float).round(2)

                # Promedio mensual Bandwidth en kHz
                bandwidth_promedios = (
                    df.groupby(["ESTACION", "FRECUENCIA (MHz)"])["Bandwidth (Hz)"]
                    .mean()
                    .reset_index()
                )
                bandwidth_promedios["Ancho de Banda (KHz)"] = (bandwidth_promedios["Bandwidth (Hz)"] / 1000).round(2)
                bandwidth_promedios = bandwidth_promedios.drop(columns=["Bandwidth (Hz)"])
                pivot = pivot.merge(bandwidth_promedios, on=["ESTACION", "FRECUENCIA (MHz)"], how="left")

                # Ordenar
                pivot = pivot.sort_values(by="FRECUENCIA (MHz)")

                # Reemplazar ceros con guiones
                pivot[todos_los_dias] = pivot[todos_los_dias].replace(0, "-")

                # Columnas manuales vacías
                pivot["Medición Manual"] = ""
                
                pivot["OBSERVACIONES"] = ""

                # Resetear índice y eliminar columna 'index' si aparece
                pivot = pivot.reset_index(drop=True)

                # Generar nombre del archivo de salida
                base_nombre = archivo_entrada.split("_")[0]
                nombre_excel = f"{base_nombre}_promediostotal.xlsx"
                ruta_excel_salida = os.path.join(ruta_salida, nombre_excel)

                # Redondear columnas numéricas
                for col in pivot.select_dtypes(include="number").columns:
                    pivot[col] = pivot[col].round(2)

                with pd.ExcelWriter(ruta_excel_salida, engine="openpyxl") as writer:
                    pivot.to_excel(writer, index=False, sheet_name="Resumen")

                

            from openpyxl import load_workbook
            from openpyxl.styles import Alignment, Font, Border, Side
            from openpyxl.utils import get_column_letter
            from openpyxl.drawing.image import Image as XLImage
            import os

            # Abrir el archivo
            wb = load_workbook(ruta_excel_salida)
            ws = wb["Resumen"] # o  si sabes el nombre de la hoja

            ws.sheet_view.showGridLines = False
            
            if base_nombre == "canar":
                base_nombre="tambo"

            # Encabezado
            encabezado_lineas = [
                "INFORME DE CONTROL TÉCNICO",
                "No. IT-CZ06-R-2025-00XX",
                "AGENCIA DE REGULACIÓN Y CONTROL DE LAS TELECOMUNICACIONES",
                "COORDINACIÓN ZONAL 6",
                "ESTACIÓN DE COMPROBACIÓN TÉCNICA",
                "FORMULARIO DE CONTROL MENSUAL DE FM",
                "CIUDAD:" + base_nombre.upper(),
                "PERIODO:" + nombre_mes_es.upper(),
                f"FECHA PRESENTACIÓN: {fecha_actual}"
            ]

            num_columnas = ws.max_column
            ultima_col = get_column_letter(num_columnas)

            # Insertar encabezado
            ws.insert_rows(1, amount=len(encabezado_lineas))
            for i, texto in enumerate(encabezado_lineas, start=1):
                celda = ws.cell(row=i, column=1, value=texto)
                ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=num_columnas)
                celda.alignment = Alignment(horizontal="center", vertical="center")
                celda.font = Font(bold=True, size=12)


            # Insertar imágenes
            ruta_imagenes = "Img"
            logo_izquierda = os.path.join(ruta_imagenes, "ARCOTEL.png")
            logo_derecha = os.path.join(ruta_imagenes, "nEcuador.png")

            if os.path.exists(logo_izquierda):
                img_left = XLImage(logo_izquierda)
                img_left.width = 500
                img_left.height = 100
                img_left.left = 10000
                img_left.top = 500
                ws.add_image(img_left, "A4")

            if os.path.exists(logo_derecha):
                img_right = XLImage(logo_derecha)
                img_right.width = 260
                img_right.height = 120
                img_right.left = 1000
                img_right.top = 500
                ws.add_image(img_right, "AH3")

            # Ajustar ancho de columnas
            for i in range(1, num_columnas + 1):
                col_letter = get_column_letter(i)
                max_length = 0
                for cell in ws[col_letter]:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[col_letter].width = max_length + 2


            # Agregar marco exterior con líneas más gruesas

            inicio_fila_tabla = len(encabezado_lineas) + 1 
            fin_fila_tabla = ws.max_row

            thin = Side(border_style="thin")
            border_thin = Border(top=thin, bottom=thin, left=thin, right=thin)
            for row in range(inicio_fila_tabla, fin_fila_tabla + 1):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).border = border_thin


            borde_grueso = Side(border_style="medium")
            for row in range(inicio_fila_tabla, fin_fila_tabla + 1):
                for col in range(1, num_columnas + 1):
                    cell = ws.cell(row=row, column=col)
                    # Obtener el borde actual (el delgado ya aplicado antes)
                    current_border = cell.border

                    # Aplicar bordes gruesos en los extremos
                    new_border = Border(
                        top=borde_grueso if row == inicio_fila_tabla else current_border.top,
                        bottom=borde_grueso if row == fin_fila_tabla else current_border.bottom,
                        left=borde_grueso if col == 1 else current_border.left,
                        right=borde_grueso if col == num_columnas else current_border.right,
                    )

                    cell.border = new_border


            alineacion_centrada = Alignment(horizontal="center", vertical="center")

            for row in range(inicio_fila_tabla, fin_fila_tabla + 1):
                for col in range(1, num_columnas + 1):
                    ws.cell(row=row, column=col).alignment = alineacion_centrada


            # Cambiar nombre del encabezado de columna y permitir salto de línea
            for col in range(1, num_columnas + 1):
                celda = ws.cell(row=inicio_fila_tabla, column=col)
                if celda.value == "Promedio(dBuV/m)":
                    celda.value = "Promedio\n(dBuV/m)"
                    celda.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

            # Cambiar nombre del encabezado de columna y permitir salto de línea
            for col in range(1, num_columnas + 1):
                celda = ws.cell(row=inicio_fila_tabla, column=col)
                if celda.value == "Ancho de Banda (KHz)":
                    celda.value = "Ancho de Banda\n(KHz)"
                    celda.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            

            # Cambiar nombre del encabezado de columna y permitir salto de línea
            for col in range(1, num_columnas + 1):
                celda = ws.cell(row=inicio_fila_tabla, column=col)
                if celda.value == "Medición Manual":
                    celda.value = "Medición Manual\nAB(KHz) o NIVEL\n(dBµV/m)"
                    celda.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            
            
            ws.row_dimensions[inicio_fila_tabla].height = 45  # o el valor que se vea bien

            # Guardar
            wb.save(ruta_excel_salida)

            print(f"✅ Archivo generado en formato Excel: {ruta_excel_salida}")

        except Exception as e:
            print(f"❌ Error procesando {archivo_entrada}: {e}")