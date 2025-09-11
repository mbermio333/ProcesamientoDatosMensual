# -*- coding: utf-8 -*-
import os
import json
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# Colores para el formato
ROJO = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
VERDE = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
AMARILLO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GRIS = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# Agrega estos umbrales TV en la sección de configuración
UMBRAL_TV_BANDA_I_III = 45  # para Bandas I-III (VHF)
UMBRAL_TV_BANDA_III = 60    # para Banda III (VHF) 
UMBRAL_TV_BANDA_IV_V = 60   # para Bandas IV-V (UHF)

# ------------------ CONFIGURACIÓN GENERAL ------------------
# Cargar configuración desde archivo
CONFIG_FILE = "config.json"

def cargar_configuracion():
    """Cargar configuración desde archivo JSON"""
    config_default = {
        "fm_path": "MedicionesFmCSV",
        "tv_path": "MedicionesTvCSV", 
        "output_path": "ReportesOcupacion",
        "emisoras_por_ciudad": {}
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Asegurar que exista la clave emisoras_por_ciudad
                if "emisoras_por_ciudad" not in config:
                    config["emisoras_por_ciudad"] = {}
                return config
        except Exception as e:
            print(f"Error cargando configuración: {e}")
            return config_default
    
    return config_default

# Cargar configuración al inicio
config = cargar_configuracion()
ruta_fm = config.get("fm_path", "MedicionesFmCSV")
ruta_tv = config.get("tv_path", "MedicionesTvCSV")
ruta_salida = config.get("output_path", "ReportesOcupacion")
fecha_actual = datetime.now().strftime("%d/%m/%Y")

# ------------------ FUNCIONES AUXILIARES ------------------

def crear_tabla_ocupacion_fm(ws, datos, umbral=60):
    """
    Crea la tabla de ocupación FM a partir de la columna J con los criterios corregidos
    """
    # Obtener los datos de la hoja
    fila_inicio = 2  # Asumiendo que la fila 1 son encabezados
    
    # Calcular estadísticas con criterios corregidos
    total_frecuencias = len(datos)
    
    # Contar frecuencias operando mayor al umbral (SOLO las que tienen Level > umbral)
    frecuencias_mayor_umbral = 0
    for fila in range(fila_inicio, ws.max_row + 1):
        nivel_celda = ws.cell(row=fila, column=5)  # Columna E = Level (dBµV/m)
        if nivel_celda.value and isinstance(nivel_celda.value, (int, float)):
            if nivel_celda.value > umbral:
                frecuencias_mayor_umbral += 1
    
    # Contar frecuencias con criterios corregidos
    frecuencias_autorizadas = 0
    frecuencias_no_autorizadas = 0
    frecuencias_observacion = 0
    frecuencias_libres = 0
    
    for fila in range(fila_inicio, ws.max_row + 1):
        estacion_celda = ws.cell(row=fila, column=2)  # Columna B = Estación
        ocupacion_celda = ws.cell(row=fila, column=4)  # Columna D = Ocupación (%)
        
        # Convertir ocupación a número
        ocupacion_valor = 0
        if ocupacion_celda.value is not None:
            try:
                ocupacion_valor = float(ocupacion_celda.value)
            except (ValueError, TypeError):
                ocupacion_valor = 0
        
        # Verificar si tiene nombre en ESTACIÓN
        tiene_nombre = estacion_celda.value and estacion_celda.value != "No identificada" and estacion_celda.value != ""
        
        if tiene_nombre:
            # Tiene nombre -> Verificar si es autorizada o no autorizada
            estacion_str = str(estacion_celda.value).lower()
            if "no autorizado" in estacion_str or "no autorizada" in estacion_str or "sis no autori" in estacion_str or "no aut" in estacion_str:
                frecuencias_no_autorizadas += 1
                # Pintar de rojo
                ws.cell(row=fila, column=1).fill = ROJO  # Frecuencia
                ws.cell(row=fila, column=2).fill = ROJO  # Estación
                ws.cell(row=fila, column=4).fill = ROJO  # Ocupación
            else:
                frecuencias_autorizadas += 1
                # Pintar de verde (aunque tenga 0% de ocupación)
                ws.cell(row=fila, column=1).fill = VERDE  # Frecuencia
                ws.cell(row=fila, column=2).fill = VERDE  # Estación
                ws.cell(row=fila, column=4).fill = VERDE  # Ocupación
        else:
            # No tiene nombre -> Verificar si es libre o en observación
            if ocupacion_valor == 0:
                frecuencias_libres += 1
            else:
                frecuencias_observacion += 1
                # Pintar de amarillo las de observación
                ws.cell(row=fila, column=1).fill = AMARILLO  # Frecuencia
                ws.cell(row=fila, column=2).fill = AMARILLO  # Estación
                ws.cell(row=fila, column=4).fill = AMARILLO  # Ocupación
    
    # VERIFICACIÓN: La suma debe coincidir con el total
    suma_categorias = (frecuencias_autorizadas + frecuencias_no_autorizadas + 
                       frecuencias_observacion + frecuencias_libres)
    
    if suma_categorias != total_frecuencias:
        print(f"⚠️  Advertencia: Suma de categorías ({suma_categorias}) no coincide con total ({total_frecuencias})")
    
    # Calcular porcentajes CORREGIDOS según los nuevos criterios
    porcentaje_ocupadas = (frecuencias_mayor_umbral / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_libres = (frecuencias_libres / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_autorizadas = (frecuencias_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_no_autorizadas = (frecuencias_no_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_observacion = (frecuencias_observacion / total_frecuencias * 100) if total_frecuencias > 0 else 0
    
    # Crear la tabla a partir de la columna J (columna 10)
    col_inicio = 10
    fila_inicio_tabla = 1
    
    # Estilos
    font_bold = Font(bold=True)
    font_normal = Font()
    alignment_center = Alignment(horizontal="center", vertical="center")
    alignment_left = Alignment(horizontal="left", vertical="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Título
    ws.cell(row=fila_inicio_tabla, column=col_inicio, value="OCUPACIÓN FM")
    ws.merge_cells(start_row=fila_inicio_tabla, start_column=col_inicio, end_row=fila_inicio_tabla, end_column=col_inicio + 3)
    titulo_cell = ws.cell(row=fila_inicio_tabla, column=col_inicio)
    titulo_cell.font = Font(bold=True, size=14)
    titulo_cell.alignment = alignment_center
    titulo_cell.fill = GRIS
    
    # Encabezados de la tabla
    datos_tabla = [
        ["UMBRAL", umbral, "% FRECUENCIAS OCUPADAS", f"{porcentaje_ocupadas:.2f}%"],
        ["TOTAL DE FRECUENCIAS MONITOREADAS", total_frecuencias, "% FRECUENCIAS LIBRES", f"{porcentaje_libres:.2f}%"],
        ["FRECUENCIAS OPERANDO MAYOR AL UMBRAL", frecuencias_mayor_umbral, "% AUTORIZADAS", f"{porcentaje_autorizadas:.2f}%"],
        ["FRECUENCIAS AUTORIZADAS", frecuencias_autorizadas, "% NO AUTORIZADAS", f"{porcentaje_no_autorizadas:.2f}%"],
        ["FRECUENCIAS EN OBSERVACIÓN", frecuencias_observacion, "% INTERMODULACIÓN O RUIDO", f"{porcentaje_observacion:.2f}%"],
        ["FRECUENCIAS NO AUTORIZADOS", frecuencias_no_autorizadas, "", ""],
        ["FRECUENCIAS LIBRES", frecuencias_libres, "", ""]
    ]
    
    # Escribir datos de la tabla
    for i, fila_datos in enumerate(datos_tabla, start=fila_inicio_tabla + 1):
        # Columna J: Descripción
        celda_j = ws.cell(row=i, column=col_inicio, value=fila_datos[0])
        celda_j.font = font_bold
        celda_j.alignment = alignment_left
        celda_j.border = thin_border
        
        # Colorear solo las celdas de la columna J (encabezados)
        if "FRECUENCIAS AUTORIZADAS" in fila_datos[0]:
            celda_j.fill = VERDE
        elif "FRECUENCIAS EN OBSERVACIÓN" in fila_datos[0]:
            celda_j.fill = AMARILLO
        elif "FRECUENCIAS NO AUTORIZADOS" in fila_datos[0]:
            celda_j.fill = ROJO
        elif "FRECUENCIAS LIBRES" in fila_datos[0]:
            celda_j.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        
        # Columna K: Valor numérico
        if fila_datos[1] != "":
            celda_k = ws.cell(row=i, column=col_inicio + 1, value=fila_datos[1])
            celda_k.font = font_bold
            celda_k.alignment = alignment_center
            celda_k.border = thin_border
        
        # Columna L: Encabezado de porcentaje
        if fila_datos[2] != "":
            celda_l = ws.cell(row=i, column=col_inicio + 2, value=fila_datos[2])
            celda_l.font = font_bold
            celda_l.alignment = alignment_left
            celda_l.border = thin_border
        
        # Columna M: Valor de porcentaje
        if fila_datos[3] != "":
            celda_m = ws.cell(row=i, column=col_inicio + 3, value=fila_datos[3])
            celda_m.font = font_bold
            celda_m.alignment = alignment_center
            celda_m.border = thin_border
    
    # Ajustar anchos de columnas
    anchos_columnas = [35, 15, 25, 15]  # J, K, L, M
    for i, ancho in enumerate(anchos_columnas, start=col_inicio):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    
    return {
        "total_frecuencias": total_frecuencias,
        "frecuencias_mayor_umbral": frecuencias_mayor_umbral,
        "frecuencias_autorizadas": frecuencias_autorizadas,
        "frecuencias_no_autorizadas": frecuencias_no_autorizadas,
        "frecuencias_observacion": frecuencias_observacion,
        "frecuencias_libres": frecuencias_libres,
        "porcentajes": {
            "ocupadas": porcentaje_ocupadas,
            "libres": porcentaje_libres,
            "autorizadas": porcentaje_autorizadas,
            "no_autorizadas": porcentaje_no_autorizadas,
            "observacion": porcentaje_observacion
        }
    }

def crear_tablas_ocupacion_tv(ws, datos):
    """
    Crea las tablas de ocupación TV por bandas a partir de la columna K
    """
    # Definir umbrales por banda
    umbrales_por_banda = {
        "Bandas I-III (VHF)": UMBRAL_TV_BANDA_I_III,
        "Banda III (VHF)": UMBRAL_TV_BANDA_III,
        "Bandas IV-V (UHF)": UMBRAL_TV_BANDA_IV_V
    }
    
    # Obtener los datos de la hoja
    fila_inicio = 2  # Asumiendo que la fila 1 son encabezados
    
    # Separar datos por banda
    datos_por_banda = {}
    for banda in umbrales_por_banda.keys():
        datos_por_banda[banda] = []
    
    # Recorrer todas las filas y agrupar por banda
    for fila in range(fila_inicio, ws.max_row + 1):
        banda_celda = ws.cell(row=fila, column=3)  # Columna C = Banda
        if banda_celda.value in umbrales_por_banda:
            datos_por_banda[banda_celda.value].append(fila)
    
    # Crear tabla para cada banda
    col_inicio = 11  # Columna K
    fila_actual = 1
    separacion_entre_tablas = 2
    
    for banda, filas_banda in datos_por_banda.items():
        if not filas_banda:
            continue  # Saltar bandas sin datos
        
        umbral = umbrales_por_banda[banda]
        
        # Calcular estadísticas para esta banda
        total_frecuencias = len(filas_banda)
        
        # Contar frecuencias operando mayor al umbral
        frecuencias_mayor_umbral = 0
        for fila in filas_banda:
            nivel_celda = ws.cell(row=fila, column=6)  # Columna F = Level (dBµV/m)
            if nivel_celda.value and isinstance(nivel_celda.value, (int, float)):
                if nivel_celda.value > umbral:
                    frecuencias_mayor_umbral += 1
        
        # Contar frecuencias con criterios específicos
        frecuencias_autorizadas = 0
        frecuencias_no_autorizadas = 0
        frecuencias_observacion = 0
        frecuencias_libres = 0
        
        for fila in filas_banda:
            estacion_celda = ws.cell(row=fila, column=2)  # Columna B = Estación
            ocupacion_celda = ws.cell(row=fila, column=5)  # Columna E = Ocupación (%)
            
            # Convertir ocupación a número
            ocupacion_valor = 0
            if ocupacion_celda.value is not None:
                try:
                    ocupacion_valor = float(ocupacion_celda.value)
                except (ValueError, TypeError):
                    ocupacion_valor = 0
            
            # Verificar si tiene nombre en ESTACIÓN
            tiene_nombre = estacion_celda.value and estacion_celda.value != "No identificada" and estacion_celda.value != ""
            
            if tiene_nombre:
                # Tiene nombre -> Verificar si es autorizada o no autorizada
                estacion_str = str(estacion_celda.value).lower()
                if any(x in estacion_str for x in ["no autorizado", "no autorizada", "no aut"]):
                    frecuencias_no_autorizadas += 1
                    # Pintar de rojo
                    ws.cell(row=fila, column=1).fill = ROJO  # Frecuencia
                    ws.cell(row=fila, column=2).fill = ROJO  # Estación
                    ws.cell(row=fila, column=5).fill = ROJO  # Ocupación
                else:
                    frecuencias_autorizadas += 1
                    # Pintar de verde
                    ws.cell(row=fila, column=1).fill = VERDE  # Frecuencia
                    ws.cell(row=fila, column=2).fill = VERDE  # Estación
                    ws.cell(row=fila, column=5).fill = VERDE  # Ocupación
            else:
                # No tiene nombre -> Verificar si es libre o en observación
                if ocupacion_valor == 0:
                    frecuencias_libres += 1
                else:
                    frecuencias_observacion += 1
                    # Pintar de amarillo
                    ws.cell(row=fila, column=1).fill = AMARILLO  # Frecuencia
                    ws.cell(row=fila, column=2).fill = AMARILLO  # Estación
                    ws.cell(row=fila, column=5).fill = AMARILLO  # Ocupación
        
        # Calcular porcentajes
        porcentaje_ocupadas = (frecuencias_mayor_umbral / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_libres = (frecuencias_libres / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_autorizadas = (frecuencias_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_no_autorizadas = (frecuencias_no_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_observacion = (frecuencias_observacion / total_frecuencias * 100) if total_frecuencias > 0 else 0
        
        # Estilos
        font_bold = Font(bold=True)
        alignment_center = Alignment(horizontal="center", vertical="center")
        alignment_left = Alignment(horizontal="left", vertical="center")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Título de la banda
        titulo_banda = f"OCUPACIÓN {banda.upper()}"
        ws.cell(row=fila_actual, column=col_inicio, value=titulo_banda)
        ws.merge_cells(start_row=fila_actual, start_column=col_inicio, end_row=fila_actual, end_column=col_inicio + 3)
        titulo_cell = ws.cell(row=fila_actual, column=col_inicio)
        titulo_cell.font = Font(bold=True, size=14)
        titulo_cell.alignment = alignment_center
        titulo_cell.fill = GRIS
        
        # Encabezados de la tabla
        datos_tabla = [
            ["UMBRAL", umbral, "% FRECUENCIAS OCUPADAS", f"{porcentaje_ocupadas:.2f}%"],
            ["TOTAL DE FRECUENCIAS MONITOREADAS", total_frecuencias, "% FRECUENCIAS LIBRES", f"{porcentaje_libres:.2f}%"],
            ["FRECUENCIAS OPERANDO MAYOR AL UMBRAL", frecuencias_mayor_umbral, "% AUTORIZADAS", f"{porcentaje_autorizadas:.2f}%"],
            ["FRECUENCIAS AUTORIZADAS", frecuencias_autorizadas, "% NO AUTORIZADAS", f"{porcentaje_no_autorizadas:.2f}%"],
            ["FRECUENCIAS EN OBSERVACIÓN", frecuencias_observacion, "% INTERMODULACIÓN O RUIDO", f"{porcentaje_observacion:.2f}%"],
            ["FRECUENCIAS NO AUTORIZADOS", frecuencias_no_autorizadas, "", ""],
            ["FRECUENCIAS LIBRES", frecuencias_libres, "", ""]
        ]
        
        # Escribir datos de la tabla
        for i, fila_datos in enumerate(datos_tabla, start=fila_actual + 1):
            # Columna K: Descripción
            celda_k = ws.cell(row=i, column=col_inicio, value=fila_datos[0])
            celda_k.font = font_bold
            celda_k.alignment = alignment_left
            celda_k.border = thin_border
            
            # Colorear solo las celdas de la columna K (encabezados)
            if "FRECUENCIAS AUTORIZADAS" in fila_datos[0]:
                celda_k.fill = VERDE
            elif "FRECUENCIAS EN OBSERVACIÓN" in fila_datos[0]:
                celda_k.fill = AMARILLO
            elif "FRECUENCIAS NO AUTORIZADOS" in fila_datos[0]:
                celda_k.fill = ROJO
            elif "FRECUENCIAS LIBRES" in fila_datos[0]:
                celda_k.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            
            # Columna L: Valor numérico
            if fila_datos[1] != "":
                celda_l = ws.cell(row=i, column=col_inicio + 1, value=fila_datos[1])
                celda_l.font = font_bold
                celda_l.alignment = alignment_center
                celda_l.border = thin_border
            
            # Columna M: Encabezado de porcentaje
            if fila_datos[2] != "":
                celda_m = ws.cell(row=i, column=col_inicio + 2, value=fila_datos[2])
                celda_m.font = font_bold
                celda_m.alignment = alignment_left
                celda_m.border = thin_border
            
            # Columna N: Valor de porcentaje
            if fila_datos[3] != "":
                celda_n = ws.cell(row=i, column=col_inicio + 3, value=fila_datos[3])
                celda_n.font = font_bold
                celda_n.alignment = alignment_center
                celda_n.border = thin_border
        
        # Ajustar anchos de columnas para esta tabla
        anchos_columnas = [35, 15, 25, 15]  # K, L, M, N
        for i, ancho in enumerate(anchos_columnas, start=col_inicio):
            ws.column_dimensions[get_column_letter(i)].width = ancho
        
        # Actualizar fila actual para la próxima tabla
        fila_actual = i + separacion_entre_tablas + 1
        
        print(f"✅ Tabla creada para {banda}: {total_frecuencias} frecuencias")
    
    return True


def buscar_emisora_por_frecuencia(ciudad, frecuencia, tipo, tolerancia=0.1):
    """
    Busca una emisora por frecuencia en una ciudad específica
    tolerancia: margen de error en MHz para coincidir frecuencias
    """
    config = cargar_configuracion()
    emisoras_por_ciudad = config.get("emisoras_por_ciudad", {})
    
    # Normalizar nombre de ciudad (manejar diferentes representaciones de "cañar")
    ciudad_normalizada = ciudad.lower().strip()
    
    # Buscar coincidencias para "cañar" en diferentes representaciones
    posibles_nombres_canar = ["cañar", "cañar", "canar", "caÃ±ar"]
    if any(nombre in ciudad_normalizada for nombre in posibles_nombres_canar):
        # Buscar la clave exacta en el config.json
        claves_config = list(emisoras_por_ciudad.keys())
        clave_canar = None
        for clave in claves_config:
            clave_normalizada = clave.lower().strip()
            if any(nombre in clave_normalizada for nombre in posibles_nombres_canar):
                clave_canar = clave
                break
        
        if clave_canar:
            ciudad_normalizada = clave_canar
        else:
            ciudad_normalizada = "cañar"
    
    if ciudad_normalizada not in emisoras_por_ciudad:
        print(f"⚠️  Ciudad '{ciudad}' no encontrada en config.json")
        print(f"Ciudades disponibles: {list(emisoras_por_ciudad.keys())}")
        # Intentar buscar por similitud
        for clave_real in emisoras_por_ciudad.keys():
            if ciudad.lower() in clave_real.lower() or clave_real.lower() in ciudad.lower():
                ciudad_normalizada = clave_real
                print(f"✅ Usando ciudad similar: {clave_real}")
                break
        else:
            return None
    
    emisoras = emisoras_por_ciudad[ciudad_normalizada].get(tipo, [])
    
    if not emisoras:
        print(f"⚠️  No hay emisoras de tipo {tipo} para la ciudad {ciudad_normalizada}")
        return None
    
    for emisora in emisoras:
        try:
            freq_emisora = float(emisora.get("frecuencia", 0))
            if abs(freq_emisora - frecuencia) <= tolerancia:
                return emisora.get("nombre", "Desconocido")
        except (ValueError, TypeError):
            continue
    
    print(f"⚠️  No se encontró emisora para frecuencia {frecuencia} MHz en {ciudad_normalizada} (tolerancia: {tolerancia} MHz)")
    return None

def obtener_emisoras_ciudad(ciudad):
    """Obtiene todas las emisoras de una ciudad"""
    config = cargar_configuracion()
    return config.get("emisoras_por_ciudad", {}).get(ciudad, {"FM": [], "TV": []})





def inicializar_directorios():
    """Crear los directorios necesarios si no existen"""
    os.makedirs(ruta_salida, exist_ok=True)
    os.makedirs(ruta_fm, exist_ok=True)
    os.makedirs(ruta_tv, exist_ok=True)

def obtener_codigo_base(base):
    """Obtiene el código correspondiente según el nombre de la base"""
    correspondencia = {
        "zamora": "SCS-L01",
        "loja": "SCS-L02", 
        "cañar": "SCS-L03",  # ñ normal
        "cañar": "SCS-L03",  # ñ con tilde combinable (n + ˜)
        "macas": "SCS-L04",
        "machala": "SCC-L04",
        "cuenca": "SCS-L05"
    }
    
    # Normalizar el nombre de la base
    base_normalizada = base.lower().strip()
    
    # Manejar diferentes representaciones de "cañar"
    if (base_normalizada == "cañar" or 
        base_normalizada == "cañar" or  # ñ con tilde combinable
        base_normalizada == "canar" or   # sin tilde
        base_normalizada == "caÃ±ar"):   # posible encoding issue
        base_normalizada = "cañar"
    
    return correspondencia.get(base_normalizada, f"SCS-{base.upper()}")

def obtener_nombre_mes_es(numero_mes):
    """Convierte el número de mes a nombre en español"""
    meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    return meses[numero_mes - 1] if 1 <= numero_mes <= 12 else ""

def obtener_base(nombre_archivo):
    """Obtiene el nombre base del archivo"""
    #print(nombre_archivo)
    return nombre_archivo.split("_")[0].lower().strip()

def reducir_archivo_csv(ruta_archivo, tipo):
    """
    Reduce el tamaño del archivo CSV conservando solo las filas necesarias
    según el tipo (FM o TV) y sobreescribe el archivo original
    """
    try:
        # Leer el archivo completo
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Convertir frecuencia a MHz
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        
        # Filtrar según el tipo
        if tipo == "FM":
            # Para FM: rango desde 88.1 MHz, conservar primeras 95 filas
            df_filtrado = df[(df["Frecuencia (MHz)"] >= 88.1)]
            if len(df_filtrado) > 95:
                df_filtrado = df_filtrado.head(95)
        else:  # TV
            # Para TV: rango 55.25-693.25 MHz, conservar primeras 45 filas
            df_filtrado = df[(df["Frecuencia (MHz)"] >= 55.25) & (df["Frecuencia (MHz)"] <= 693.25)]
            if len(df_filtrado) > 45:
                df_filtrado = df_filtrado.head(45)
        
        # Eliminar la columna temporal
        if "Frecuencia (MHz)" in df_filtrado.columns:
            df_filtrado = df_filtrado.drop(columns=["Frecuencia (MHz)"])
        
        # Guardar el archivo reducido (sobreescribir el original)
        df_filtrado.to_csv(ruta_archivo, index=False, encoding='latin-1')
        
        return True
        
    except Exception as e:
        print(f"Error reduciendo archivo {ruta_archivo}: {e}")
        return False

def formatear_hoja_ocupacion(ws, datos, tipo, ciudad):
    """Formatea una hoja de ocupación con bordes y estilos, incluyendo columna Estación"""
    # Limpiar hoja existente
    ws.delete_rows(1, ws.max_row)
    
    # Agregar encabezado según el tipo
    if tipo == "FM":
        encabezados = [
            "Frecuencia (MHz)", "Estación", "FECHA DE SUSCRIPCION", "Ocupación (%)",
            "Level (dBµV/m)", "Bandwidth (Hz)", "Offset (Hz)", "FM (kHz)"
        ]
    else:  # TV
        encabezados = [
            "Frecuencia (MHz)", "Estación", "Banda", "Canal", "Ocupación (%)",
            "Level (dBµV/m)", "Bandwidth (Hz)", "Offset (Hz)", "AM (%)"
        ]
    
    # Escribir encabezados
    for col, encabezado in enumerate(encabezados, 1):
        ws.cell(row=1, column=col, value=encabezado)
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")
    
    # Escribir datos con columna Estación
    for fila_idx, (_, fila) in enumerate(datos.iterrows(), 2):
        frecuencia = fila["Frecuencia (MHz)"]
        
        # Buscar emisora por frecuencia
        nombre_emisora = buscar_emisora_por_frecuencia(ciudad, frecuencia, tipo)
        
        if tipo == "FM":
            ws.cell(row=fila_idx, column=1, value=frecuencia)
            ws.cell(row=fila_idx, column=2, value=nombre_emisora or "No identificada")
            ws.cell(row=fila_idx, column=3, value=fila["FECHA DE SUSCRIPCION"])
            ws.cell(row=fila_idx, column=4, value=fila["Ocupación (%)"])
            ws.cell(row=fila_idx, column=5, value=fila["Level (dBµV/m)"])
            ws.cell(row=fila_idx, column=6, value=fila["Bandwidth (Hz)"])
            ws.cell(row=fila_idx, column=7, value=fila["Offset (Hz)"])
            ws.cell(row=fila_idx, column=8, value=fila["FM (kHz)"])
        else:  # TV
            ws.cell(row=fila_idx, column=1, value=frecuencia)
            ws.cell(row=fila_idx, column=2, value=nombre_emisora or "No identificada")
            ws.cell(row=fila_idx, column=3, value=fila["Banda"])
            ws.cell(row=fila_idx, column=4, value=fila["Canal"])
            ws.cell(row=fila_idx, column=5, value=fila["Ocupación (%)"])
            ws.cell(row=fila_idx, column=6, value=fila["Level (dBµV/m)"])
            ws.cell(row=fila_idx, column=7, value=fila["Bandwidth (Hz)"])
            ws.cell(row=fila_idx, column=8, value=fila["Offset (Hz)"])
            ws.cell(row=fila_idx, column=9, value=fila["AM (%)"])
    
    # Aplicar bordes y formato a la tabla principal
    thin = Side(border_style="thin")
    borde_grueso = Side(border_style="medium")
    
    num_columnas = len(encabezados)
    num_filas = len(datos) + 1
    
    for row in range(1, num_filas + 1):
        for col in range(1, num_columnas + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            if row == 1:  # Encabezados
                cell.border = Border(top=borde_grueso, bottom=thin,
                                   left=borde_grueso if col == 1 else thin,
                                   right=borde_grueso if col == num_columnas else thin)
            elif row == num_filas:  # Última fila
                cell.border = Border(top=thin, bottom=borde_grueso,
                                   left=borde_grueso if col == 1 else thin,
                                   right=borde_grueso if col == num_columnas else thin)
            elif col == 1:  # Primera columna
                cell.border = Border(left=borde_grueso, top=thin, bottom=thin, right=thin)
            elif col == num_columnas:  # Última columna
                cell.border = Border(right=borde_grueso, top=thin, bottom=thin, left=thin)
    
    # Ajustar anchos de columnas de la tabla principal
    for col in range(1, num_columnas + 1):
        col_letter = get_column_letter(col)
        max_length = 0
        for cell in ws[col_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        # Ancho especial para la columna Estación
        if col == 2:  # Columna Estación
            ws.column_dimensions[col_letter].width = max(max_length + 2, 25)
        else:
            ws.column_dimensions[col_letter].width = max_length + 2
    
    # Solo para FM, crear la tabla de ocupación
    if tipo == "FM":
        crear_tabla_ocupacion_fm(ws, datos)

    if tipo == "TV":
        crear_tablas_ocupacion_tv(ws, datos)

def limpiar_valor_numerico(valor):
    """Limpia y convierte valores numéricos, manejando formatos con coma decimal"""
    if pd.isna(valor) or valor is None:
        return np.nan
    
    # Convertir a string y limpiar
    str_valor = str(valor).strip()
    
    # Manejar valores inválidos
    if str_valor == '-1e+040' or 'nan' in str_valor.lower():
        return np.nan
    
    # Reemplazar coma por punto para decimales
    str_valor = str_valor.replace(',', '.')
    
    # Eliminar espacios y caracteres no numéricos (excepto punto y signo negativo)
    str_valor = ''.join(c for c in str_valor if c.isdigit() or c in ['.', '-'])
    
    try:
        return float(str_valor)
    except ValueError:
        return np.nan

def buscar_columna_por_patron(df, patrones):
    """Busca una columna en el DataFrame que coincida con alguno de los patrones"""
    for patron in patrones:
        for col in df.columns:
            if patron.lower() in col.lower():
                return col
    return None

def obtener_banda_por_frecuencia(freq):
    """Determina la banda según la frecuencia"""
    # Bandas de TV según estándares internacionales
    if 55.25 <= freq <= 88: 
        return "Bandas I-III (VHF)"
    elif 174 <= freq <= 216: 
        return "Banda III (VHF)"
    elif 470 <= freq <= 608: 
        return "Bandas IV-V (UHF)"
    elif 614 <= freq <= 698: 
        return "Bandas IV-V (UHF)"
    else: 
        return "Otra banda"

def procesar_archivo_fm(ruta_archivo, base):
    """Procesa archivo FM y extrae datos de ocupación en el rango desde 88.1 MHz"""
    try:
        # Primero reducir el archivo
        if not reducir_archivo_csv(ruta_archivo, "FM"):
            return None, None
            
        # Leer el archivo reducido
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Filtrar por rango de frecuencia (desde 88.1 MHz)
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        df_filtrado = df[(df["Frecuencia (MHz)"] >= 88.1)]
        
        if df_filtrado.empty:
            return None, None
        
        # Obtener mes de los datos - manejar diferentes formatos de fecha
        try:
            df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date
        except:
            try:
                df["Tiempo"] = pd.to_datetime(df["Tiempo"], errors='coerce').dt.date
            except:
                df["Tiempo"] = None
        
        # Usar la fecha del sistema si no se puede determinar del archivo
        if df["Tiempo"].isna().all():
            mes_objetivo = datetime.now().month
        else:
            mes_objetivo = df["Tiempo"].dropna().apply(lambda x: x.month).value_counts().idxmax()
        
        # Buscar las columnas necesarias
        columna_ocupacion = buscar_columna_por_patron(df_filtrado, ['ocupaci'])
        columna_level = buscar_columna_por_patron(df_filtrado, ['level', 'nivel'])
        columna_bandwidth = buscar_columna_por_patron(df_filtrado, ['bandwidth', 'ancho de banda'])
        columna_offset = buscar_columna_por_patron(df_filtrado, ['offset', 'desplazamiento'])
        columna_fm = buscar_columna_por_patron(df_filtrado, ['fm', 'frecuencia modulada'])
        
        # Si no se encuentra FM, buscar AM y convertir a kHz
        if columna_fm is None:
            columna_am = buscar_columna_por_patron(df_filtrado, ['am', 'amplitud modulada'])
            if columna_am is not None:
                columna_fm = columna_am
        
        # Calcular ocupación usando los valores reales del archivo
        ocupacion_data = []
        
        if columna_ocupacion:
            # Usar los valores reales de ocupación del archivo
            for _, row in df_filtrado.iterrows():
                try:
                    # Limpiar y convertir el valor de ocupación
                    ocupacion_val = limpiar_valor_numerico(row[columna_ocupacion])
                    if not np.isnan(ocupacion_val):
                        # Obtener los demás valores
                        level_val = limpiar_valor_numerico(row[columna_level]) if columna_level else np.nan
                        bandwidth_val = limpiar_valor_numerico(row[columna_bandwidth]) if columna_bandwidth else np.nan
                        offset_val = limpiar_valor_numerico(row[columna_offset]) if columna_offset else np.nan
                        
                        # Para FM, si se encontró AM en lugar de FM, convertir a kHz
                        fm_val = np.nan
                        if columna_fm:
                            fm_val = limpiar_valor_numerico(row[columna_fm])
                            # Si el valor es de AM (%), convertirlo a FM (kHz)
                            # Asumimos que valores mayores a 100 son kHz, menores son %
                            if fm_val <= 100:
                                fm_val = fm_val * 10  # Convertir % a kHz (aproximación)
                        
                        ocupacion_data.append({
                            "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                            "FECHA DE SUSCRIPCION": datetime.now().strftime("%Y-%m-%d"),
                            "Ocupación (%)": ocupacion_val,  # Valor original con decimales
                            "Level (dBµV/m)": level_val,
                            "Bandwidth (Hz)": bandwidth_val,
                            "Offset (Hz)": offset_val,
                            "FM (kHz)": fm_val
                        })
                except (ValueError, TypeError):
                    continue
        else:
            # Si no hay columna de ocupación, usar 100% para frecuencias con señal
            for _, row in df_filtrado.iterrows():
                try:
                    # Obtener los demás valores
                    level_val = limpiar_valor_numerico(row[columna_level]) if columna_level else np.nan
                    bandwidth_val = limpiar_valor_numerico(row[columna_bandwidth]) if columna_bandwidth else np.nan
                    offset_val = limpiar_valor_numerico(row[columna_offset]) if columna_offset else np.nan
                    
                    # Para FM, si se encontró AM en lugar de FM, convertir to kHz
                    fm_val = np.nan
                    if columna_fm:
                        fm_val = limpiar_valor_numerico(row[columna_fm])
                        # Si el valor es de AM (%), convertirlo a FM (kHz)
                        if fm_val <= 100:
                            fm_val = fm_val * 10  # Convertir % a kHz (aproximación)
                    
                    ocupacion_data.append({
                        "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                        "FECHA DE SUSCRIPCION": datetime.now().strftime("%Y-%m-%d"),
                        "Ocupación (%)": 100.0,
                        "Level (dBµV/m)": level_val,
                        "Bandwidth (Hz)": bandwidth_val,
                        "Offset (Hz)": offset_val,
                        "FM (kHz)": fm_val
                    })
                except (ValueError, TypeError):
                    continue
        
        resultado = pd.DataFrame(ocupacion_data)
        resultado["Mes"] = mes_objetivo
        return resultado, base
        
    except Exception as e:
        print(f"Error procesando archivo FM {ruta_archivo}: {e}")
        import traceback
        traceback.print_exc()
        return None, base

def procesar_archivo_tv(ruta_archivo, base):
    """Procesa archivo TV y extrae datos de ocupación en el rango 55.25-693.25 MHz"""
    try:
        # Primero reducir el archivo
        if not reducir_archivo_csv(ruta_archivo, "TV"):
            return None, None
            
        # Leer el archivo reducido
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Filtrar por rango de frecuencia correcto (55.25-693.25 MHz)
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        df_filtrado = df[(df["Frecuencia (MHz)"] >= 55.25) & (df["Frecuencia (MHz)"] <= 693.25)]
        
        if df_filtrado.empty:
            return None, None
        
        # Obtener mes de los datos - manejar diferentes formatos de fecha
        try:
            df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date
        except:
            try:
                df["Tiempo"] = pd.to_datetime(df["Tiempo"], errors='coerce').dt.date
            except:
                df["Tiempo"] = None
        
        # Usar la fecha del sistema si no se puede determinar del archivo
        if df["Tiempo"].isna().all():
            mes_objetivo = datetime.now().month
        else:
            mes_objetivo = df["Tiempo"].dropna().apply(lambda x: x.month).value_counts().idxmax()
        
        # Buscar las columnas necesarias
        columna_ocupacion = buscar_columna_por_patron(df_filtrado, ['ocupaci'])
        columna_canal = buscar_columna_por_patron(df_filtrado, ['canal', 'channel'])
        columna_level = buscar_columna_por_patron(df_filtrado, ['level', 'nivel'])
        columna_bandwidth = buscar_columna_por_patron(df_filtrado, ['bandwidth', 'ancho de banda'])
        columna_offset = buscar_columna_por_patron(df_filtrado, ['offset', 'desplazamiento'])
        columna_am = buscar_columna_por_patron(df_filtrado, ['am', 'amplitud modulada'])
        
        # Calcular ocupación usando los valores reales del archivo
        ocupacion_data = []
        
        if columna_ocupacion:
            # Usar los valores reales de ocupación del archivo
            for _, row in df_filtrado.iterrows():
                try:
                    # Limpiar y convertir el valor de ocupación
                    ocupacion_val = limpiar_valor_numerico(row[columna_ocupacion])
                    if not np.isnan(ocupacion_val):
                        # Obtener los demás valores
                        canal_val = row[columna_canal] if columna_canal else np.nan
                        level_val = limpiar_valor_numerico(row[columna_level]) if columna_level else np.nan
                        bandwidth_val = limpiar_valor_numerico(row[columna_bandwidth]) if columna_bandwidth else np.nan
                        offset_val = limpiar_valor_numerico(row[columna_offset]) if columna_offset else np.nan
                        am_val = limpiar_valor_numerico(row[columna_am]) if columna_am else np.nan
                        
                        # Determinar la banda según la frecuencia
                        banda_val = obtener_banda_por_frecuencia(row["Frecuencia (MHz)"])
                        
                        ocupacion_data.append({
                            "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                            "Banda": banda_val,
                            "Canal": canal_val,  # Valor original del canal
                            "Ocupación (%)": ocupacion_val,  # Valor original con decimales
                            "Level (dBµV/m)": level_val,
                            "Bandwidth (Hz)": bandwidth_val,
                            "Offset (Hz)": offset_val,
                            "AM (%)": am_val
                        })
                except (ValueError, TypeError):
                    continue
        else:
            # Si no hay columna de ocupación, usar 100% para frecuencias con señal
            for _, row in df_filtrado.iterrows():
                try:
                    # Obtener los demás valores
                    canal_val = row[columna_canal] if columna_canal else np.nan
                    level_val = limpiar_valor_numerico(row[columna_level]) if columna_level else np.nan
                    bandwidth_val = limpiar_valor_numerico(row[columna_bandwidth]) if columna_bandwidth else np.nan
                    offset_val = limpiar_valor_numerico(row[columna_offset]) if columna_offset else np.nan
                    am_val = limpiar_valor_numerico(row[columna_am]) if columna_am else np.nan
                    
                    # Determinar la banda según la frecuencia
                    banda_val = obtener_banda_por_frecuencia(row["Frecuencia (MHz)"])
                    
                    ocupacion_data.append({
                        "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                        "Banda": banda_val,
                        "Canal": canal_val,  # Valor original del canal
                        "Ocupación (%)": 100.0,
                        "Level (dBµV/m)": level_val,
                        "Bandwidth (Hz)": bandwidth_val,
                        "Offset (Hz)": offset_val,
                        "AM (%)": am_val
                    })
                except (ValueError, TypeError):
                    continue
        
        resultado = pd.DataFrame(ocupacion_data)
        resultado["Mes"] = mes_objetivo
        return resultado, base
        
    except Exception as e:
        print(f"Error procesando archivo TV {ruta_archivo}: {e}")
        import traceback
        traceback.print_exc()
        return None, base
    
def debug_emisoras_config():
    """Debug: mostrar el contenido completo de emisoras en config.json"""
    config = cargar_configuracion()
    emisoras_por_ciudad = config.get("emisoras_por_ciudad", {})
    
    print("=== DEBUG: CONTENIDO DE config.json ===")
    print(f"Número de ciudades: {len(emisoras_por_ciudad)}")
    
    for ciudad, tipos in emisoras_por_ciudad.items():
        print(f"\n--- CIUDAD: '{ciudad}' (tipo: {type(ciudad)}) ---")
        # Mostrar representación raw de la cadena
        print(f"Representación: {repr(ciudad)}")
        
        for tipo, emisoras in tipos.items():
            print(f"  {tipo}: {len(emisoras)} emisoras")
            for i, emisora in enumerate(emisoras[:5]):  # Mostrar solo las primeras 5
                nombre = emisora.get('nombre', 'Sin nombre')
                freq = emisora.get('frecuencia', 0)
                print(f"    {i+1}. {nombre} - {freq} MHz")
            if len(emisoras) > 5:
                print(f"    ... y {len(emisoras) - 5} más")

def verificar_encoding_config():
    """Verifica el encoding del archivo config.json"""
    try:
        with open(CONFIG_FILE, 'rb') as f:
            contenido = f.read()
            print(f"Encoding detectado: {contenido.decode('utf-8', errors='replace')[:100]}...")
            
        # Intentar diferentes encodings
        encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
        for encoding in encodings:
            try:
                with open(CONFIG_FILE, 'r', encoding=encoding) as f:
                    data = json.load(f)
                    print(f"✅ Encoding {encoding} funciona")
                    ciudades = list(data.get('emisoras_por_ciudad', {}).keys())
                    print(f"Ciudades encontradas con {encoding}: {ciudades}")
            except UnicodeDecodeError:
                print(f"❌ Encoding {encoding} falla")
    except Exception as e:
        print(f"Error al verificar encoding: {e}")

# ------------------ FUNCIÓN PRINCIPAL DE PROCESAMIENTO ------------------

def procesar_ocupacion(callback_progreso=None, callback_log=None):
    """
    Función principal que procesa datos de ocupación de espectro
    """
    # Inicializar directorios
    inicializar_directorios()
    
    if callback_log:
        callback_log("Iniciando procesamiento de ocupación de espectro...")
        callback_log(f"Ruta FM: {ruta_fm}")
        callback_log(f"Ruta TV: {ruta_tv}")
        callback_log(f"Ruta salida: {ruta_salida}")
    
    # Obtener listas de archivos
    try:
        archivos_fm = {obtener_base(f): os.path.join(ruta_fm, f) for f in os.listdir(ruta_fm) if f.endswith(".csv")}
        archivos_tv = {obtener_base(f): os.path.join(ruta_tv, f) for f in os.listdir(ruta_tv) if f.endswith(".csv")}
        
        if callback_log:
            callback_log(f"Encontrados {len(archivos_fm)} archivos FM y {len(archivos_tv)} archivos TV")
    except Exception as e:
        if callback_log:
            callback_log(f"Error al leer archivos: {str(e)}")
        return False

    # Encontrar bases comunes entre FM y TV
    bases_comunes = set(archivos_fm.keys()).intersection(archivos_tv.keys())
    total_bases = len(bases_comunes)
    
    if total_bases == 0:
        if callback_log:
            callback_log("No se encontraron bases comunes entre FM and TV")
        return False
    
    if callback_log:
        callback_log(f"Procesando {total_bases} bases comunes")
    
    # Procesar cada base común
    for i, base in enumerate(bases_comunes):
        if callback_progreso:
            # Calcular progreso (0-100)
            progreso = int((i / total_bases) * 100)
            callback_progreso(progreso)
            
        if callback_log:
            callback_log(f"Procesando base: {base}")
        
        try:
            # En la función procesar_ocupacion(), modifica esta parte:

            # En la función procesar_ocupacion(), modifica esta parte:

        # Procesar archivos FM y TV (ahora pasamos la base como parámetro)
            datos_fm, base_fm = procesar_archivo_fm(archivos_fm[base], base)
            datos_tv, base_tv = procesar_archivo_tv(archivos_tv[base], base)

            if datos_fm is None or datos_tv is None:
                if callback_log:
                    callback_log(f"❌ No se pudieron procesar los datos para {base}")
                continue

            # Verificar que ambos archivos sean del mismo mes
            mes_fm = datos_fm["Mes"].iloc[0] if not datos_fm.empty else None
            mes_tv = datos_tv["Mes"].iloc[0] if not datos_tv.empty else None

            if mes_fm != mes_tv:
                if callback_log:
                    callback_log(f"⚠️  Los archivos de {base} son de meses diferentes: FM={mes_fm}, TV={mes_tv}")

            # Usar el mes de FM como referencia (or TV si FM no está disponible)
            mes_referencia = mes_fm if mes_fm is not None else mes_tv

            # Crear libro de Excel
            wb = Workbook()

            # Crear hoja para FM
            if "Sheet" in wb.sheetnames:
                ws_fm = wb["Sheet"]
                ws_fm.title = "Datos FM"
            else:
                ws_fm = wb.create_sheet("Datos FM")

            # AQUÍ ESTÁ LA CORRECCIÓN - Solo llamar a formatear_hoja_ocupacion si hay datos
            if not datos_fm.empty:
                formatear_hoja_ocupacion(ws_fm, datos_fm.drop(columns=["Mes"]), "FM", base)
            else:
                if callback_log:
                    callback_log(f"⚠️  No hay datos FM para {base}")

            # Crear hoja para TV
            ws_tv = wb.create_sheet("Datos TV")
            if not datos_tv.empty:
                formatear_hoja_ocupacion(ws_tv, datos_tv.drop(columns=["Mes"]), "TV", base)
            else:
                if callback_log:
                    callback_log(f"⚠️  No hay datos TV para {base}")

            # Eliminar hoja por defecto si existe
            if "Sheet" in wb.sheetnames and wb.sheetnames[0] == "Sheet":
                del wb["Sheet"]
            
            # Generar nombre de archivo
            codigo_base = obtener_codigo_base(base)

            # En la parte donde generas el nombre del archivo:
            def normalizar_nombre_ciudad(base):
                """Normaliza el nombre de la ciudad para el nombre del archivo"""
                base_normalizada = base.lower().strip()
                
                # Manejar "cañar" y sus variantes
                if (base_normalizada == "cañar" or 
                    base_normalizada == "cañar" or 
                    base_normalizada == "canar" or 
                    base_normalizada == "caÃ±ar"):
                    return "TAMBO"
                else:
                    return base.upper()

            # Uso:
            nombre_ciudad = normalizar_nombre_ciudad(base)

            nombre_mes_completo = obtener_nombre_mes_es(mes_referencia) if mes_referencia else "Desconocido"
            nombre_salida = f"{codigo_base}_Ocupacion{nombre_ciudad}_{nombre_mes_completo}2025.xlsx"
            
            # Guardar archivo
            wb.save(os.path.join(ruta_salida, nombre_salida))
            
            if callback_log:
                callback_log(f"✅ Archivo generado: {nombre_salida}")
                
        except Exception as e:
            if callback_log:
                callback_log(f"❌ Error procesando base {base}: {str(e)}")
            import traceback
            traceback.print_exc()
    
    if callback_progreso:
        callback_progreso(100)
        
    if callback_log:
        callback_log("Procesamiento de ocupación completado")
    
    return True

# ------------------ EJECUCIÓN DIRECTA (para testing) ------------------

if __name__ == "__main__":
    # Si se ejecuta directamente, usar callbacks simples
    verificar_encoding_config()
    
    # Luego debug del contenido
    debug_emisoras_config()


    def mostrar_progreso(progreso):
        print(f"Progreso: {progreso}%")
    
    def mostrar_log(mensaje):
        print(mensaje)
    
    # Procesar datos de ocupación
    resultado = procesar_ocupacion(mostrar_progreso, mostrar_log)
    
    if resultado:
        print("Procesamiento de ocupación completado con éxito")
    else:
        print("Ocurrieron errores durante el procesamiento")