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
from openpyxl.chart import PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList

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

# Agregar estas funciones al inicio del archivo, después de las importaciones

def normalizar_nombre_ciudad(nombre):
    """Normaliza el nombre de la ciudad para consistencia"""
    if not nombre or not isinstance(nombre, str):
        return ""
    
    nombre = nombre.lower().strip()
    
    # Manejar todas las variantes de "cañar"
    #if nombre in ["cañar", "cañar", "canar", "caã±ar", "tambo"]:
    #    return "TAMBO"
    
    # Mapeo de otras ciudades si es necesario
    mapeo_ciudades = {
        "zamora": "ZAMORA",
        "loja": "LOJA", 
        "macas": "MACAS",
        "tambo":"TAMBO",
        "machala": "MACHALA",
        "cuenca": "CUENCA"
    }

    return mapeo_ciudades.get(nombre, nombre.upper())

def guardar_advertencia_ocupacion_cero(datos_problematicos, ruta_salida):
    """
    Guarda las frecuencias con ocupación 0% en un archivo JSON
    """
    try:
        archivo_advertencia = os.path.join(ruta_salida, "AdvertenciaOcup.json")
        
        # Preparar datos para JSON - ACTUALIZADO PARA INCLUIR AM
        datos_json = {
            "fecha_generacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_frecuencias_problematicas": (
                len(datos_problematicos.get("FM", [])) + 
                len(datos_problematicos.get("TV", [])) +
                len(datos_problematicos.get("AM", []))  # ← NUEVO: Incluir AM
            ),
            "frecuencias_fm": datos_problematicos.get("FM", []),
            "frecuencias_tv": datos_problematicos.get("TV", []),
            "frecuencias_am": datos_problematicos.get("AM", [])  # ← NUEVO: Incluir AM
        }
        
        with open(archivo_advertencia, 'w', encoding='utf-8') as f:
            json.dump(datos_json, f, indent=4, ensure_ascii=False)
        
        return archivo_advertencia
    except Exception as e:
        print(f"Error al guardar advertencia de ocupación 0%: {e}")
        return None

def detectar_frecuencias_ocupacion_cero(ws, tipo, ciudad, resultados_estadisticas=None):
    """
    Detecta frecuencias con ocupación 0% que tienen nombre en el campo Estación
    y NO terminan en "_OBSERVACION"
    Retorna lista de frecuencias problemáticas
    """
    frecuencias_problematicas = []
    
    try:
        # Determinar columnas según el tipo
        if tipo == "FM":
            col_frecuencia = 1  # Columna A
            col_estacion = 2    # Columna B
            col_ocupacion = 4   # Columna D
            col_level = 5       # Columna E
        else:  # TV
            col_frecuencia = 1  # Columna A
            col_estacion = 2    # Columna B
            col_ocupacion = 5   # Columna E (para TV)
            col_level = 6       # Columna F (para TV)
            col_banda = 3       # Columna C (para TV)
        
        # Recorrer filas (empezando desde la fila 2, ya que la 1 es encabezado)
        for fila in range(2, ws.max_row + 1):
            try:
                # Obtener valores de las celdas
                frecuencia_celda = ws.cell(row=fila, column=col_frecuencia)
                estacion_celda = ws.cell(row=fila, column=col_estacion)
                ocupacion_celda = ws.cell(row=fila, column=col_ocupacion)
                level_celda = ws.cell(row=fila, column=col_level)
                
                # Verificar que todas las celdas tengan valores
                if (frecuencia_celda.value is None or 
                    estacion_celda.value is None or 
                    ocupacion_celda.value is None):
                    continue
                
                # Convertir ocupación a número
                try:
                    ocupacion_valor = float(ocupacion_celda.value)
                except (ValueError, TypeError):
                    continue
                
                # Verificar si es frecuencia problemática: ocupación 0% y tiene nombre de estación
                # que NO termina en "_OBSERVACION"
                estacion_nombre = str(estacion_celda.value).strip()
                tiene_nombre_valido = (estacion_nombre and 
                                     estacion_nombre != "No identificada" and 
                                     estacion_nombre != "" and
                                     not estacion_nombre.upper().endswith("_OBSERVACION"))  # NUEVA CONDICIÓN
                
                if ocupacion_valor == 0 and tiene_nombre_valido:
                    # Obtener información adicional
                    frecuencia_valor = frecuencia_celda.value
                    level_valor = level_celda.value if level_celda.value is not None else "N/A"
                    
                    # Para TV, obtener la banda
                    tipo_detallado = tipo
                    if tipo == "TV":
                        banda_celda = ws.cell(row=fila, column=col_banda)
                        banda_valor = banda_celda.value if banda_celda.value else "Desconocida"
                        tipo_detallado = f"TV-{banda_valor}"
                    
                    # Crear registro de frecuencia problemática
                    frecuencia_problematica = {
                        "ciudad": normalizar_nombre_ciudad(ciudad),
                        "estacion": estacion_nombre,
                        "tipo": tipo_detallado,
                        "frecuencia": frecuencia_valor,
                        "ocupacion": ocupacion_valor,
                        "level": level_valor,
                        "fila_excel": fila
                    }
                    
                    frecuencias_problematicas.append(frecuencia_problematica)
                    
            except Exception as e:
                print(f"Error procesando fila {fila} en {tipo}: {e}")
                continue
                
    except Exception as e:
        print(f"Error general detectando frecuencias con ocupación 0% en {tipo}: {e}")
    
    return frecuencias_problematicas

def construir_nombre_completo_spectra(info_spectra, nombre_original):
    """
    Construye el nombre completo usando la información de SPECTRA
    Formato: "RED (NOMBRES) - ESTACION"
    """
    if not info_spectra:
        return nombre_original
    
    red = info_spectra.get("RED", "").strip()
    nombres = info_spectra.get("NOMBRES", "").strip()
    estacion = info_spectra.get("ESTACION", "").strip()
    
    # Si no hay información adicional útil, mantener el nombre original
    if not red and not nombres and not estacion:
        return nombre_original
    
    partes = []
    
    # Agregar RED si existe
    if red:
        partes.append(red)
    
    # Agregar NOMBRES entre paréntesis si es diferente de RED
    if nombres and nombres != red:
        if partes:  # Si ya hay RED, poner NOMBRES entre paréntesis
            partes[-1] = f"{partes[-1]} ({nombres})"
        else:
            partes.append(nombres)
    
    # Agregar ESTACION si existe y es diferente
    if estacion and estacion not in nombres and estacion not in red:
        partes.append(estacion)
    
    # Si no hay partes después del procesamiento, mantener original
    if not partes:
        return nombre_original
    
    nombre_completo = " - ".join(partes)
    
    # Limitar la longitud si es muy largo
    #if len(nombre_completo) > 100:
    #    nombre_completo = nombre_completo[:97] + "..."
    
    return nombre_completo



def insertar_umbrales_excel(archivo_excel, umbrales_ciudad):
    """
    Inserta los umbrales en las posiciones específicas del archivo Excel
    """
    try:
        import openpyxl
        
        # Abrir el archivo Excel
        workbook = openpyxl.load_workbook(archivo_excel)
        
        # Insertar umbrales en las posiciones especificadas
        if 'Datos FM' in workbook.sheetnames:
            hoja_fm = workbook['Datos FM']
            # Umbral FM en K2
            hoja_fm['K2'] = umbrales_ciudad['FM']

        if 'Datos AM' in workbook.sheetnames:
            hoja_fm = workbook['Datos AM']
            # Umbral AM en K2
            hoja_fm['K2'] = umbrales_ciudad['AM']
        
        if 'Datos TV' in workbook.sheetnames:
            hoja_tv = workbook['Datos TV']
            
            # Obtener configuración de TV
            tv_config = umbrales_ciudad['TV']
            
            if tv_config['tipo'] == 'general':
                # Umbral general para todas las bandas
                umbral_general = tv_config['valor']
                hoja_tv['L2'] = umbral_general   # Banda I-III
                hoja_tv['L12'] = umbral_general  # Banda III
                hoja_tv['L22'] = umbral_general  # Banda IV-V
            else:
                # Umbrales específicos por banda
                valores_bandas = tv_config['valores']
                hoja_tv['L2'] = valores_bandas.get('Banda I-III', 47.0)   # Banda I-III
                hoja_tv['L12'] = valores_bandas.get('Banda III', 56.0)     # Banda III
                hoja_tv['L22'] = valores_bandas.get('Banda IV-V', 64.0)    # Banda IV-V
        
        # Guardar los cambios
        workbook.save(archivo_excel)
        workbook.close()
        
        return True
        
    except Exception as e:
        print(f"Error al insertar umbrales en {archivo_excel}: {e}")
        return False



###############----------------SPECTRA---------------####################
def extraer_ciudad_desde_nombre_archivo(nombre_archivo):
    """
    Extrae el nombre de la ciudad desde el nombre del archivo Excel
    Ejemplo: 'SCS-L01_OcupacionZamora_Octubre2025.xlsx' -> 'zamora'
    """
    nombre_archivo = nombre_archivo.lower()
    
    # Mapeo de códigos a ciudades
    codigo_ciudad_map = {
        "scs-l01": "zamora",
        "scs-l02": "loja", 
        "scs-l03": "tambo",
        "scs-l04": "macas",
        "scc-l04": "machala",
        "scs-l05": "cuenca"
    }
    
    # Buscar por código primero
    for codigo, ciudad in codigo_ciudad_map.items():
        if codigo in nombre_archivo:
            print(f"🔍 Ciudad detectada por código {codigo}: {ciudad}")
            return ciudad
    
    # Buscar por nombre de ciudad directamente en el patrón _Ocupacion[Ciudad]_
    import re
    patron_ocupacion = re.search(r'_ocupacion([a-z]+)_', nombre_archivo)
    if patron_ocupacion:
        ciudad_encontrada = patron_ocupacion.group(1)
        print(f"🔍 Ciudad detectada por patrón Ocupacion: {ciudad_encontrada}")
        return ciudad_encontrada
    
    # Buscar por nombre de ciudad directamente
    ciudades_conocidas = ["zamora", "loja", "macas", "tambo", "machala", "cuenca"]
    for ciudad in ciudades_conocidas:
        if ciudad in nombre_archivo:
            print(f"🔍 Ciudad detectada por nombre directo: {ciudad}")
            return ciudad
    
    print(f"⚠️  No se pudo detectar ciudad desde: {nombre_archivo}")
    return "desconocida"


def obtener_ciudades_spectra_disponibles():
    """
    Obtiene la lista de ciudades disponibles en el directorio SPECTRA_Filtrado
    """
    try:
        directorio_spectra = "SPECTRA_Filtrado"
        if not os.path.exists(directorio_spectra):
            print(f"⚠️  Directorio {directorio_spectra} no existe")
            return []
        
        archivos_spectra = [f for f in os.listdir(directorio_spectra) if f.endswith('_spectra.json')]
        ciudades = [f.replace('_spectra.json', '') for f in archivos_spectra]
        
        print(f"🏙️  Ciudades disponibles en SPECTRA: {ciudades}")
        return ciudades
    except Exception as e:
        print(f"❌ Error al listar ciudades SPECTRA: {e}")
        return []











def cargar_datos_spectra(ciudad):
    """
    Carga los datos adicionales del archivo spectra JSON para una ciudad específica
    con manejo mejorado de errores
    """
    try:
        # Normalizar nombre de ciudad
        ciudad = ciudad.lower().strip()
        ruta_spectra = os.path.join("SPECTRA_Filtrado", f"{ciudad}_spectra.json")
        
        if not os.path.exists(ruta_spectra):
            # Intentar con variaciones del nombre
            variaciones = [
                ciudad,
                ciudad.replace(' ', '_'),
                ciudad.replace('-', '_'),
                ciudad.title(),
                ciudad.upper()
            ]
            
            for variacion in variaciones:
                ruta_alternativa = os.path.join("SPECTRA_Filtrado", f"{variacion}_spectra.json")
                if os.path.exists(ruta_alternativa):
                    ruta_spectra = ruta_alternativa
                    print(f"✅ Encontrado archivo spectra con variación: {variacion}")
                    break
        
        if not os.path.exists(ruta_spectra):
            print(f"⚠️  No se encontró archivo spectra para {ciudad}: {ruta_spectra}")
            # Mostrar qué archivos sí existen
            directorio = "SPECTRA_Filtrado"
            if os.path.exists(directorio):
                archivos = [f for f in os.listdir(directorio) if f.endswith('_spectra.json')]
                print(f"   Archivos disponibles: {archivos}")
            return []
        
        with open(ruta_spectra, 'r', encoding='utf-8') as f:
            datos_spectra = json.load(f)
        
        print(f"✅ Datos spectra cargados para {ciudad}: {len(datos_spectra)} registros")
        return datos_spectra
        
    except Exception as e:
        print(f"❌ Error cargando datos spectra para {ciudad}: {e}")
        return []
    
def verificar_archivos_spectra():
    """
    Verifica qué archivos spectra están disponibles
    """
    directorio = "SPECTRA_Filtrado"
    if not os.path.exists(directorio):
        print(f"❌ El directorio {directorio} no existe")
        return
    
    archivos = [f for f in os.listdir(directorio) if f.endswith('_spectra.json')]
    print("=== ARCHIVOS SPECTRA DISPONIBLES ===")
    for archivo in archivos:
        ciudad = archivo.replace('_spectra.json', '')
        ruta_completa = os.path.join(directorio, archivo)
        tamaño = os.path.getsize(ruta_completa)
        print(f"📍 {ciudad}: {archivo} ({tamaño} bytes)")
    
def buscar_info_spectra_por_frecuencia(datos_spectra, frecuencia, tipo, tolerancia=0.1):
    """
    Busca información adicional en los datos spectra por frecuencia y tipo
    con los mismos criterios de tolerancia que buscar_emisora_por_frecuencia
    """
    if not datos_spectra:
        return None
    
    # Mapear tipos internos a tipos del spectra
    tipo_spectra_map = {
        "FM": "FM - Frecuencia Modulada",
        "AM": "AM - Amplitud Modulada", 
        "TV": "TV - Televisión Abierta"
    }
    
    tipo_buscar = tipo_spectra_map.get(tipo, tipo)
    
    # CASO ESPECÍFICO PARA AM - TOLERANCIAS MÁS ESTRICTAS (igual que buscar_emisora_por_frecuencia)
    if tipo == "AM":
        # Para AM: coincidencia exacta muy estricta primero
        for registro in datos_spectra:
            registro_servicio = registro.get("SERVICIO", "")
            registro_frecuencia = registro.get("FRECUENCIA", 0)
            
            if registro_servicio == tipo_buscar:
                try:
                    if abs(registro_frecuencia - frecuencia) <= 0.001:  # Coincidencia muy exacta (0.001 MHz)
                        return registro
                except (ValueError, TypeError):
                    continue
        
        # Luego buscar con tolerancia muy reducida
        for registro in datos_spectra:
            registro_servicio = registro.get("SERVICIO", "")
            registro_frecuencia = registro.get("FRECUENCIA", 0)
            
            if registro_servicio == tipo_buscar:
                try:
                    if abs(registro_frecuencia - frecuencia) <= 0.005:  # Tolerancia muy reducida (0.005 MHz)
                        return registro
                except (ValueError, TypeError):
                    continue
        
        return None
    
    else:
        # PARA FM Y TV - MANTENER TOLERANCIAS ORIGINALES
        # Buscar coincidencia exacta primero
        for registro in datos_spectra:
            registro_servicio = registro.get("SERVICIO", "")
            registro_frecuencia = registro.get("FRECUENCIA", 0)
            
            if registro_servicio == tipo_buscar:
                try:
                    if abs(registro_frecuencia - frecuencia) <= 0.01:  # Coincidencia exacta (0.01 MHz)
                        return registro
                except (ValueError, TypeError):
                    continue
        
        # Luego buscar con tolerancia normal
        for registro in datos_spectra:
            registro_servicio = registro.get("SERVICIO", "")
            registro_frecuencia = registro.get("FRECUENCIA", 0)
            
            if registro_servicio == tipo_buscar:
                try:
                    if abs(registro_frecuencia - frecuencia) <= tolerancia:  # Tolerancia normal (0.1 MHz por defecto)
                        return registro
                except (ValueError, TypeError):
                    continue
    
    return None

def enriquecer_nombre_estacion(nombre_original, info_spectra):
    """
    Enriquece el nombre de la estación con información adicional del spectra
    Formato: "RED (NOMBRES) - ESTACION"
    """
    if not info_spectra:
        return nombre_original
    
    red = info_spectra.get("RED", "").strip()
    nombres = info_spectra.get("NOMBRES", "").strip()
    estacion = info_spectra.get("ESTACION", "").strip()
    
    # Si no hay información adicional útil, mantener el nombre original
    if not red and not nombres and not estacion:
        return nombre_original
    
    # Si el nombre original ya contiene esta información, no duplicar
    nombre_original_lower = nombre_original.lower()
    if (red and red.lower() in nombre_original_lower) or \
       (nombres and nombres.lower() in nombre_original_lower) or \
       (estacion and estacion.lower() in nombre_original_lower):
        return nombre_original
    
    partes = []
    
    # Agregar RED si existe y es diferente de nombres
    if red and red != nombres:
        partes.append(red)
    
    # Agregar NOMBRES entre paréntesis si es diferente de RED
    if nombres and nombres != red:
        if partes:  # Si ya hay RED, poner NOMBRES entre paréntesis
            partes[-1] = f"{partes[-1]} ({nombres})"
        else:
            partes.append(nombres)
    
    # Agregar ESTACION si existe y es diferente
    if estacion and estacion not in nombres and estacion not in red:
        partes.append(estacion)
    
    # Si no hay partes después del procesamiento, mantener original
    if not partes:
        return nombre_original
    
    nombre_enriquecido = " - ".join(partes)
    
    # Limitar la longitud si es muy largo
    #if len(nombre_enriquecido) > 100:
    #    nombre_enriquecido = nombre_enriquecido[:97] + "..."
    
    return nombre_enriquecido

def crear_hoja_datos_manual(wb, nombre_archivo_excel=None):
    """
    Crea la hoja 'DATOS Manual' con resumen de frecuencias autorizadas, no autorizadas
    y en observación para FM, TV y AM, extrayendo datos de las hojas existentes.
    Ahora enriquece los nombres con información adicional del spectra.
    """
    # Verificar si las hojas de datos existen
    hojas_existentes = [hoja for hoja in ["Datos FM", "Datos TV", "Datos AM"] if hoja in wb.sheetnames]
    if not hojas_existentes:
        print("⚠️  No se encontraron hojas de datos")
        return wb
    
    # CORRECCIÓN: Inicializar ciudad primero
    ciudad = "desconocida"  # <-- INICIALIZAR AQUÍ

    # Opción 1: Desde el nombre del archivo Excel (más confiable)
    if nombre_archivo_excel:
        ciudad = extraer_ciudad_desde_nombre_archivo(nombre_archivo_excel)
        print(f"🔍 Ciudad detectada desde nombre archivo: {ciudad}")
    
    # Opción 2: Buscar en las hojas (como respaldo)
    if ciudad == "desconocida":
        for nombre_hoja in wb.sheetnames:
            if "Datos FM" in nombre_hoja or "Datos TV" in nombre_hoja or "Datos AM" in nombre_hoja:
                try:
                    ws_temp = wb[nombre_hoja]
                    for fila in range(2, min(10, ws_temp.max_row + 1)):
                        estacion_celda = ws_temp.cell(row=fila, column=2)
                        if estacion_celda.value:
                            ciudades_conocidas = ["cuenca", "zamora", "loja", "macas", "tambo", "machala"]
                            for ciudad_temp in ciudades_conocidas:
                                if ciudad_temp in str(estacion_celda.value).lower():
                                    ciudad = ciudad_temp
                                    break
                        if ciudad != "desconocida":
                            break
                except:
                    pass
    
    # Opción 3: Mostrar ciudades disponibles y usar la primera
    if ciudad == "desconocida":
        ciudades_disponibles = obtener_ciudades_spectra_disponibles()
        if ciudades_disponibles:
            ciudad = ciudades_disponibles[0]  # Usar la primera disponible
            print(f"⚠️  Usando primera ciudad disponible: {ciudad}")
        else:
            print("❌ No se pudo determinar la ciudad y no hay archivos SPECTRA disponibles")
            return wb
    
    print(f"🔍 Ciudad final para enriquecimiento: {ciudad}")


    # Cargar datos spectra para la ciudad
    datos_spectra = cargar_datos_spectra(ciudad)
    print(f"🔍 Cargados {len(datos_spectra)} registros spectra para {ciudad}")
    


    # Debug: mostrar algunos registros spectra para TV
    if datos_spectra:
        print("📺 Muestra de registros SPECTRA para TV:")
        tv_registros = [r for r in datos_spectra if r.get("SERVICIO") == "TV - Televisión Abierta"]
        for i, registro in enumerate(tv_registros[:5]):
            print(f"   {i+1}. Frec: {registro.get('FRECUENCIA')} - RED: '{registro.get('RED')}' - NOMBRES: '{registro.get('NOMBRES')}' - ESTACION: '{registro.get('ESTACION')}'")
    
    # Crear o limpiar hoja existente
    if "DATOS Manual" in wb.sheetnames:
        ws_manual = wb["DATOS Manual"]
        ws_manual.delete_rows(1, ws_manual.max_row)
        ws_manual.delete_cols(1, ws_manual.max_column)
    else:
        ws_manual = wb.create_sheet("DATOS Manual")
    
    # Definir estilos y colores (mantener igual)
    color_titulo_principal = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    color_subtitulo = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
    color_encabezado = PatternFill(start_color="95B3D7", end_color="95B3D7", fill_type="solid")
    color_autorizada = PatternFill(start_color="C4D79B", end_color="C4D79B", fill_type="solid")
    color_no_autorizada = PatternFill(start_color="FCD5B5", end_color="FCD5B5", fill_type="solid")
    color_observacion = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    fuente_titulo = Font(bold=True, size=14)
    fuente_subtitulo = Font(bold=True, size=12)
    fuente_encabezado = Font(bold=True, color="FFFFFF")
    fuente_normal = Font(size=11)
    fuente_resumen = Font(bold=True, size=11)
    
    borde_fino = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    alineacion_centro = Alignment(horizontal='center', vertical='center')
    alineacion_izquierda = Alignment(horizontal='left', vertical='center')
    
    # Título principal
    ws_manual.merge_cells('A1:K1')
    celda_titulo = ws_manual['A1']
    celda_titulo.value = "DATOS DEL MONITOREO MANUAL"
    celda_titulo.fill = color_titulo_principal
    celda_titulo.font = fuente_titulo
    celda_titulo.alignment = alineacion_centro
    
    # Subtítulos para cada tipo
    ws_manual.merge_cells('A2:C2')
    celda_fm = ws_manual['A2']
    celda_fm.value = "FRECUENCIA MODULADA"
    celda_fm.fill = color_subtitulo
    celda_fm.font = fuente_subtitulo
    celda_fm.alignment = alineacion_centro
    
    ws_manual.merge_cells('E2:G2')
    celda_am = ws_manual['E2']
    celda_am.value = "AMPLITUD MODULADA"
    celda_am.fill = color_subtitulo
    celda_am.font = fuente_subtitulo
    celda_am.alignment = alineacion_centro
    
    ws_manual.merge_cells('I2:K2')
    celda_tv = ws_manual['I2']
    celda_tv.value = "TELEVISION ABIERTA"
    celda_tv.fill = color_subtitulo
    celda_tv.font = fuente_subtitulo
    celda_tv.alignment = alineacion_centro
    
    # Obtener datos de FM (MODIFICADO: incluir frecuencia para el enriquecimiento)
    datos_fm_autorizadas = []
    datos_fm_no_autorizadas = []
    datos_fm_observacion = []
    
    if "Datos FM" in wb.sheetnames:
        ws_fm = wb["Datos FM"]
        for fila in range(2, ws_fm.max_row + 1):
            estacion = ws_fm.cell(row=fila, column=2).value
            frecuencia = ws_fm.cell(row=fila, column=1).value
            color_celda = ws_fm.cell(row=fila, column=2).fill
            
            estacion_original = estacion
        
            if estacion and frecuencia:
                # Buscar información adicional en spectra
                info_spectra = buscar_info_spectra_por_frecuencia(datos_spectra, frecuencia, "FM")
                
                # DEBUG: Mostrar información de spectra encontrada
                if info_spectra:
                    print(f"✅ SPECTRA encontrado para {frecuencia} MHz:")
                    print(f"   RED: '{info_spectra.get('RED', '')}'")
                    print(f"   NOMBRES: '{info_spectra.get('NOMBRES', '')}'")
                    print(f"   ESTACION: '{info_spectra.get('ESTACION', '')}'")
                
                # Construir nombre enriquecido con el formato deseado
                nombre_enriquecido = construir_nombre_completo_spectra(info_spectra, estacion_original)
                
                # DEBUG: Comparar nombres
                if nombre_enriquecido != estacion_original:
                    print(f"🔄 Nombre enriquecido: '{estacion_original}' -> '{nombre_enriquecido}'")
                
                # Clasificar según el color (usando el nombre original para la clasificación)
                if color_celda.start_color.index == VERDE.start_color.index:
                    datos_fm_autorizadas.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                elif color_celda.start_color.index == ROJO.start_color.index:
                    datos_fm_no_autorizadas.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                elif color_celda.start_color.index == AMARILLO.start_color.index:
                    datos_fm_observacion.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
    
    # Obtener datos de AM (MODIFICADO: incluir frecuencia para el enriquecimiento)
    datos_am_autorizadas = []
    datos_am_no_autorizadas = []
    datos_am_observacion = []
    
    if "Datos AM" in wb.sheetnames:
        ws_am = wb["Datos AM"]
        for fila in range(2, ws_am.max_row + 1):
            estacion = ws_am.cell(row=fila, column=2).value
            frecuencia = ws_am.cell(row=fila, column=1).value
            color_celda = ws_am.cell(row=fila, column=2).fill
            
            estacion_original = estacion
        
            if estacion and frecuencia:
                # Buscar información adicional en spectra
                info_spectra = buscar_info_spectra_por_frecuencia(datos_spectra, frecuencia, "AM")
                
                # DEBUG: Mostrar información de spectra encontrada
                if info_spectra:
                    print(f"✅ SPECTRA encontrado para {frecuencia} MHz:")
                    print(f"   RED: '{info_spectra.get('RED', '')}'")
                    print(f"   NOMBRES: '{info_spectra.get('NOMBRES', '')}'")
                    print(f"   ESTACION: '{info_spectra.get('ESTACION', '')}'")
                
                # Construir nombre enriquecido con el formato deseado
                nombre_enriquecido = construir_nombre_completo_spectra(info_spectra, estacion_original)
                
                # DEBUG: Comparar nombres
                if nombre_enriquecido != estacion_original:
                    print(f"🔄 Nombre enriquecido: '{estacion_original}' -> '{nombre_enriquecido}'")
                
                # Clasificar según el color (usando el nombre original para la clasificación)
                if color_celda.start_color.index == VERDE.start_color.index:
                    datos_am_autorizadas.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                elif color_celda.start_color.index == ROJO.start_color.index:
                    datos_am_no_autorizadas.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                elif color_celda.start_color.index == AMARILLO.start_color.index:
                    datos_am_observacion.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido

    
    # Obtener datos de TV (MODIFICADO: incluir frecuencia para el enriquecimiento)
    datos_tv_autorizadas = []
    datos_tv_no_autorizadas = []
    datos_tv_observacion = []
    # Debug inicial de SPECTRA
    print(f"=== DEBUG SPECTRA PARA {ciudad} ===")
    print(f"Total registros SPECTRA: {len(datos_spectra)}")
    tv_spectra = [r for r in datos_spectra if r.get("SERVICIO") == "TV - Televisión Abierta"]
    print(f"Registros TV en SPECTRA: {len(tv_spectra)}")
    for i, reg in enumerate(tv_spectra[:3]):
        print(f"  {i+1}. Frec: {reg.get('FRECUENCIA')} - RED: '{reg.get('RED')}'")
    if "Datos TV" in wb.sheetnames:
        ws_tv = wb["Datos TV"]
        for fila in range(2, ws_tv.max_row + 1):
            estacion = ws_tv.cell(row=fila, column=2).value
            frecuencia = ws_tv.cell(row=fila, column=1).value
            color_celda = ws_tv.cell(row=fila, column=2).fill
            
            
                
            estacion_original = estacion
        
            if estacion and frecuencia:
                # Buscar información adicional en spectra
                info_spectra = buscar_info_spectra_por_frecuencia(datos_spectra, frecuencia, "TV")
                
                # DEBUG: Mostrar información de spectra encontrada
                if info_spectra:
                    print(f"✅ SPECTRA encontrado para {frecuencia} MHz:")
                    print(f"   RED: '{info_spectra.get('RED', '')}'")
                    print(f"   NOMBRES: '{info_spectra.get('NOMBRES', '')}'")
                    print(f"   ESTACION: '{info_spectra.get('ESTACION', '')}'")
                
                # Construir nombre enriquecido con el formato deseado
                nombre_enriquecido = construir_nombre_completo_spectra(info_spectra, estacion_original)
                
                # DEBUG: Comparar nombres
                if nombre_enriquecido != estacion_original:
                    print(f"🔄 Nombre enriquecido: '{estacion_original}' -> '{nombre_enriquecido}'")
                
                # Clasificar según el color (usando el nombre original para la clasificación)
                if color_celda.start_color.index == VERDE.start_color.index:
                    datos_tv_autorizadas.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                elif color_celda.start_color.index == ROJO.start_color.index:
                    datos_tv_no_autorizadas.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                elif color_celda.start_color.index == AMARILLO.start_color.index:
                    datos_tv_observacion.append((frecuencia, nombre_enriquecido))  # Guardar nombre enriquecido
                



    # --- SECCIÓN FM (Columnas A-C) ---
    fila_actual = 3
    
    # Resumen FM
    ws_manual[f'A{fila_actual}'] = "FRECUENCIAS AUTORIZADAS"
    ws_manual[f'B{fila_actual}'] = len(datos_fm_autorizadas)
    ws_manual[f'A{fila_actual}'].fill = color_autorizada
    ws_manual[f'B{fila_actual}'].fill = color_autorizada
    ws_manual[f'A{fila_actual}'].font = fuente_resumen
    ws_manual[f'B{fila_actual}'].font = fuente_resumen
    ws_manual[f'A{fila_actual}'].alignment = alineacion_izquierda
    ws_manual[f'B{fila_actual}'].alignment = alineacion_centro
    ws_manual[f'A{fila_actual}'].border = borde_fino
    ws_manual[f'B{fila_actual}'].border = borde_fino
    
    fila_actual += 1
    ws_manual[f'A{fila_actual}'] = "FRECUENCIAS NO AUTORIZADAS"
    ws_manual[f'B{fila_actual}'] = len(datos_fm_no_autorizadas)
    ws_manual[f'A{fila_actual}'].fill = color_no_autorizada
    ws_manual[f'B{fila_actual}'].fill = color_no_autorizada
    ws_manual[f'A{fila_actual}'].font = fuente_resumen
    ws_manual[f'B{fila_actual}'].font = fuente_resumen
    ws_manual[f'A{fila_actual}'].alignment = alineacion_izquierda
    ws_manual[f'B{fila_actual}'].alignment = alineacion_centro
    ws_manual[f'A{fila_actual}'].border = borde_fino
    ws_manual[f'B{fila_actual}'].border = borde_fino
    
    fila_actual += 1
    ws_manual[f'A{fila_actual}'] = "EN OBSERVACIÓN"
    ws_manual[f'B{fila_actual}'] = len(datos_fm_observacion)
    ws_manual[f'A{fila_actual}'].fill = color_observacion
    ws_manual[f'B{fila_actual}'].fill = color_observacion
    ws_manual[f'A{fila_actual}'].font = fuente_resumen
    ws_manual[f'B{fila_actual}'].font = fuente_resumen
    ws_manual[f'A{fila_actual}'].alignment = alineacion_izquierda
    ws_manual[f'B{fila_actual}'].alignment = alineacion_centro
    ws_manual[f'A{fila_actual}'].border = borde_fino
    ws_manual[f'B{fila_actual}'].border = borde_fino
    
    # Detalle FM
    fila_detalle_fm = fila_actual + 2
    ws_manual[f'A{fila_detalle_fm}'] = "TIPO"
    ws_manual[f'B{fila_detalle_fm}'] = "Frecuencia (MHz)"
    ws_manual[f'C{fila_detalle_fm}'] = "ESTACION"
    
    for col in ['A', 'B', 'C']:
        celda = ws_manual[f'{col}{fila_detalle_fm}']
        celda.fill = color_encabezado
        celda.font = fuente_encabezado
        celda.alignment = alineacion_centro
        celda.border = borde_fino
    
    # Insertar datos FM (ahora con nombres enriquecidos)
    fila_actual_fm = fila_detalle_fm + 1
    for frecuencia, estacion in datos_fm_autorizadas:
        ws_manual[f'A{fila_actual_fm}'] = "AUTORIZADA"
        ws_manual[f'B{fila_actual_fm}'] = frecuencia
        ws_manual[f'C{fila_actual_fm}'] = estacion
        for col in ['A', 'B', 'C']:
            celda = ws_manual[f'{col}{fila_actual_fm}']
            celda.fill = color_autorizada
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['A', 'C']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_fm += 1
    
    for frecuencia, estacion in datos_fm_no_autorizadas:
        ws_manual[f'A{fila_actual_fm}'] = "NO AUTORIZADA"
        ws_manual[f'B{fila_actual_fm}'] = frecuencia
        ws_manual[f'C{fila_actual_fm}'] = estacion
        for col in ['A', 'B', 'C']:
            celda = ws_manual[f'{col}{fila_actual_fm}']
            celda.fill = color_no_autorizada
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['A', 'C']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_fm += 1
    
    for frecuencia, estacion in datos_fm_observacion:
        ws_manual[f'A{fila_actual_fm}'] = "EN OBSERVACIÓN"
        ws_manual[f'B{fila_actual_fm}'] = frecuencia
        ws_manual[f'C{fila_actual_fm}'] = estacion
        for col in ['A', 'B', 'C']:
            celda = ws_manual[f'{col}{fila_actual_fm}']
            celda.fill = color_observacion
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['A', 'C']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_fm += 1
    
    # --- SECCIÓN AM (Columnas E-G) ---
    fila_actual_am = 3
    
    # Resumen AM
    ws_manual[f'E{fila_actual_am}'] = "FRECUENCIAS AUTORIZADAS"
    ws_manual[f'F{fila_actual_am}'] = len(datos_am_autorizadas)
    ws_manual[f'E{fila_actual_am}'].fill = color_autorizada
    ws_manual[f'F{fila_actual_am}'].fill = color_autorizada
    ws_manual[f'E{fila_actual_am}'].font = fuente_resumen
    ws_manual[f'F{fila_actual_am}'].font = fuente_resumen
    ws_manual[f'E{fila_actual_am}'].alignment = alineacion_izquierda
    ws_manual[f'F{fila_actual_am}'].alignment = alineacion_centro
    ws_manual[f'E{fila_actual_am}'].border = borde_fino
    ws_manual[f'F{fila_actual_am}'].border = borde_fino
    
    fila_actual_am += 1
    ws_manual[f'E{fila_actual_am}'] = "FRECUENCIAS NO AUTORIZADAS"
    ws_manual[f'F{fila_actual_am}'] = len(datos_am_no_autorizadas)
    ws_manual[f'E{fila_actual_am}'].fill = color_no_autorizada
    ws_manual[f'F{fila_actual_am}'].fill = color_no_autorizada
    ws_manual[f'E{fila_actual_am}'].font = fuente_resumen
    ws_manual[f'F{fila_actual_am}'].font = fuente_resumen
    ws_manual[f'E{fila_actual_am}'].alignment = alineacion_izquierda
    ws_manual[f'F{fila_actual_am}'].alignment = alineacion_centro
    ws_manual[f'E{fila_actual_am}'].border = borde_fino
    ws_manual[f'F{fila_actual_am}'].border = borde_fino
    
    fila_actual_am += 1
    ws_manual[f'E{fila_actual_am}'] = "EN OBSERVACIÓN"
    ws_manual[f'F{fila_actual_am}'] = len(datos_am_observacion)
    ws_manual[f'E{fila_actual_am}'].fill = color_observacion
    ws_manual[f'F{fila_actual_am}'].fill = color_observacion
    ws_manual[f'E{fila_actual_am}'].font = fuente_resumen
    ws_manual[f'F{fila_actual_am}'].font = fuente_resumen
    ws_manual[f'E{fila_actual_am}'].alignment = alineacion_izquierda
    ws_manual[f'F{fila_actual_am}'].alignment = alineacion_centro
    ws_manual[f'E{fila_actual_am}'].border = borde_fino
    ws_manual[f'F{fila_actual_am}'].border = borde_fino
    
    # Detalle AM
    fila_detalle_am = fila_actual_am + 2
    ws_manual[f'E{fila_detalle_am}'] = "TIPO"
    ws_manual[f'F{fila_detalle_am}'] = "Frecuencia (MHz)"
    ws_manual[f'G{fila_detalle_am}'] = "ESTACION"
    
    for col in ['E', 'F', 'G']:
        celda = ws_manual[f'{col}{fila_detalle_am}']
        celda.fill = color_encabezado
        celda.font = fuente_encabezado
        celda.alignment = alineacion_centro
        celda.border = borde_fino
    
    # Insertar datos AM (ahora con nombres enriquecidos)
    fila_actual_am_detalle = fila_detalle_am + 1
    for frecuencia, estacion in datos_am_autorizadas:
        ws_manual[f'E{fila_actual_am_detalle}'] = "AUTORIZADA"
        ws_manual[f'F{fila_actual_am_detalle}'] = frecuencia
        ws_manual[f'G{fila_actual_am_detalle}'] = estacion
        for col in ['E', 'F', 'G']:
            celda = ws_manual[f'{col}{fila_actual_am_detalle}']
            celda.fill = color_autorizada
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['E', 'G']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_am_detalle += 1
    
    for frecuencia, estacion in datos_am_no_autorizadas:
        ws_manual[f'E{fila_actual_am_detalle}'] = "NO AUTORIZADA"
        ws_manual[f'F{fila_actual_am_detalle}'] = frecuencia
        ws_manual[f'G{fila_actual_am_detalle}'] = estacion
        for col in ['E', 'F', 'G']:
            celda = ws_manual[f'{col}{fila_actual_am_detalle}']
            celda.fill = color_no_autorizada
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['E', 'G']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_am_detalle += 1
    
    for frecuencia, estacion in datos_am_observacion:
        ws_manual[f'E{fila_actual_am_detalle}'] = "EN OBSERVACIÓN"
        ws_manual[f'F{fila_actual_am_detalle}'] = frecuencia
        ws_manual[f'G{fila_actual_am_detalle}'] = estacion
        for col in ['E', 'F', 'G']:
            celda = ws_manual[f'{col}{fila_actual_am_detalle}']
            celda.fill = color_observacion
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['E', 'G']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_am_detalle += 1
    
    # --- SECCIÓN TV (Columnas I-K) ---
    fila_actual_tv = 3
    
    # Resumen TV
    ws_manual[f'I{fila_actual_tv}'] = "FRECUENCIAS AUTORIZADAS"
    ws_manual[f'J{fila_actual_tv}'] = len(datos_tv_autorizadas)
    ws_manual[f'I{fila_actual_tv}'].fill = color_autorizada
    ws_manual[f'J{fila_actual_tv}'].fill = color_autorizada
    ws_manual[f'I{fila_actual_tv}'].font = fuente_resumen
    ws_manual[f'J{fila_actual_tv}'].font = fuente_resumen
    ws_manual[f'I{fila_actual_tv}'].alignment = alineacion_izquierda
    ws_manual[f'J{fila_actual_tv}'].alignment = alineacion_centro
    ws_manual[f'I{fila_actual_tv}'].border = borde_fino
    ws_manual[f'J{fila_actual_tv}'].border = borde_fino
    
    fila_actual_tv += 1
    ws_manual[f'I{fila_actual_tv}'] = "FRECUENCIAS NO AUTORIZADAS"
    ws_manual[f'J{fila_actual_tv}'] = len(datos_tv_no_autorizadas)
    ws_manual[f'I{fila_actual_tv}'].fill = color_no_autorizada
    ws_manual[f'J{fila_actual_tv}'].fill = color_no_autorizada
    ws_manual[f'I{fila_actual_tv}'].font = fuente_resumen
    ws_manual[f'J{fila_actual_tv}'].font = fuente_resumen
    ws_manual[f'I{fila_actual_tv}'].alignment = alineacion_izquierda
    ws_manual[f'J{fila_actual_tv}'].alignment = alineacion_centro
    ws_manual[f'I{fila_actual_tv}'].border = borde_fino
    ws_manual[f'J{fila_actual_tv}'].border = borde_fino
    
    fila_actual_tv += 1
    ws_manual[f'I{fila_actual_tv}'] = "EN OBSERVACIÓN"
    ws_manual[f'J{fila_actual_tv}'] = len(datos_tv_observacion)
    ws_manual[f'I{fila_actual_tv}'].fill = color_observacion
    ws_manual[f'J{fila_actual_tv}'].fill = color_observacion
    ws_manual[f'I{fila_actual_tv}'].font = fuente_resumen
    ws_manual[f'J{fila_actual_tv}'].font = fuente_resumen
    ws_manual[f'I{fila_actual_tv}'].alignment = alineacion_izquierda
    ws_manual[f'J{fila_actual_tv}'].alignment = alineacion_centro
    ws_manual[f'I{fila_actual_tv}'].border = borde_fino
    ws_manual[f'J{fila_actual_tv}'].border = borde_fino
    
    # Detalle TV
    fila_detalle_tv = fila_actual_tv + 2
    ws_manual[f'I{fila_detalle_tv}'] = "TIPO"
    ws_manual[f'J{fila_detalle_tv}'] = "Frecuencia (MHz)"
    ws_manual[f'K{fila_detalle_tv}'] = "ESTACION"
    
    for col in ['I', 'J', 'K']:
        celda = ws_manual[f'{col}{fila_detalle_tv}']
        celda.fill = color_encabezado
        celda.font = fuente_encabezado
        celda.alignment = alineacion_centro
        celda.border = borde_fino
    
    # Insertar datos TV (ahora con nombres enriquecidos)
    fila_actual_tv_detalle = fila_detalle_tv + 1
    for frecuencia, estacion in datos_tv_autorizadas:
        ws_manual[f'I{fila_actual_tv_detalle}'] = "AUTORIZADA"
        ws_manual[f'J{fila_actual_tv_detalle}'] = frecuencia
        ws_manual[f'K{fila_actual_tv_detalle}'] = estacion
        for col in ['I', 'J', 'K']:
            celda = ws_manual[f'{col}{fila_actual_tv_detalle}']
            celda.fill = color_autorizada
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['I', 'K']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_tv_detalle += 1
    
    for frecuencia, estacion in datos_tv_no_autorizadas:
        ws_manual[f'I{fila_actual_tv_detalle}'] = "NO AUTORIZADA"
        ws_manual[f'J{fila_actual_tv_detalle}'] = frecuencia
        ws_manual[f'K{fila_actual_tv_detalle}'] = estacion
        for col in ['I', 'J', 'K']:
            celda = ws_manual[f'{col}{fila_actual_tv_detalle}']
            celda.fill = color_no_autorizada
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['I', 'K']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_tv_detalle += 1
    
    for frecuencia, estacion in datos_tv_observacion:
        ws_manual[f'I{fila_actual_tv_detalle}'] = "EN OBSERVACIÓN"
        ws_manual[f'J{fila_actual_tv_detalle}'] = frecuencia
        ws_manual[f'K{fila_actual_tv_detalle}'] = estacion
        for col in ['I', 'J', 'K']:
            celda = ws_manual[f'{col}{fila_actual_tv_detalle}']
            celda.fill = color_observacion
            celda.font = fuente_normal
            celda.border = borde_fino
            if col in ['I', 'K']:
                celda.alignment = alineacion_izquierda
            else:
                celda.alignment = alineacion_centro
        fila_actual_tv_detalle += 1
    
    # Ajustar anchos de columnas (aumentar ancho de columna ESTACION para nombres más largos)
    ws_manual.column_dimensions['A'].width = 20
    ws_manual.column_dimensions['B'].width = 15
    ws_manual.column_dimensions['C'].width = 40  # Aumentado para nombres enriquecidos
    ws_manual.column_dimensions['E'].width = 20
    ws_manual.column_dimensions['F'].width = 15
    ws_manual.column_dimensions['G'].width = 40  # Aumentado para nombres enriquecidos
    ws_manual.column_dimensions['I'].width = 20
    ws_manual.column_dimensions['J'].width = 15
    ws_manual.column_dimensions['K'].width = 40  # Aumentado para nombres enriquecidos
    
    print(f"✅ Hoja 'DATOS Manual' creada con nombres enriquecidos para {ciudad}")
    
    return wb


def crear_grafico_pastel_openpyxl(ws, porcentaje_ocupadas, porcentaje_libres, titulo, celda_destino, identificador_unico):
    """
    Crea un gráfico de pastel directamente con openpyxl
    """
    # Crear datos para el gráfico en celdas ocultas con posición única
    fila_inicio = 100 + (identificador_unico * 10)  # Espacio suficiente entre gráficos
    col_datos = 20

    # Limpiar celdas previas (por si acaso)
    for i in range(4):  # Limpiar 4 filas
        ws.cell(row=fila_inicio + i, column=col_datos, value="")
        ws.cell(row=fila_inicio + i, column=col_datos + 1, value="")

    ws.cell(row=fila_inicio, column=col_datos, value="Ocupadas")
    ws.cell(row=fila_inicio, column=col_datos + 1, value=porcentaje_ocupadas)

    ws.cell(row=fila_inicio + 1, column=col_datos, value="Libres")
    ws.cell(row=fila_inicio + 1, column=col_datos + 1, value=porcentaje_libres)

    # Crear gráfico de pastel
    chart = PieChart()
    chart.title = titulo

    # Referencias a datos y etiquetas
    labels = Reference(ws, min_col=col_datos, min_row=fila_inicio, max_row=fila_inicio + 1)
    data = Reference(ws, min_col=col_datos + 1, min_row=fila_inicio, max_row=fila_inicio + 1)

    # Añadir datos y categorías correctamente
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(labels)

    # Configurar etiquetas de datos
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCategoryName = True
    chart.dataLabels.showVal = False
    chart.dataLabels.showSerName = False

    # Añadir gráfico a la hoja
    ws.add_chart(chart, celda_destino)

    return chart

# Agregar esta nueva función después de la función crear_grafico_pastel_openpyxl existente

def crear_grafico_pastel_estatus_openpyxl(ws, porcentaje_autorizadas, porcentaje_no_autorizadas, porcentaje_observacion, 
                                        titulo, celda_destino, identificador_unico):
    """
    Crea un gráfico de pastel para los estatus (autorizadas, no autorizadas, observación)
    """
    # Crear datos para el gráfico en celdas ocultas con posición única
    fila_inicio = 150 + (identificador_unico * 10)  # Espacio suficiente entre gráficos (diferente del anterior)
    col_datos = 25  # Columna Y para datos temporales

    # Limpiar celdas previas (por si acaso)
    for i in range(4):  # Limpiar 4 filas
        ws.cell(row=fila_inicio + i, column=col_datos, value="")
        ws.cell(row=fila_inicio + i, column=col_datos + 1, value="")

    # Crear datos para el gráfico
    ws.cell(row=fila_inicio, column=col_datos, value="AUTORIZADAS")
    ws.cell(row=fila_inicio, column=col_datos + 1, value=porcentaje_autorizadas)

    ws.cell(row=fila_inicio + 1, column=col_datos, value="NO AUTORIZADAS")
    ws.cell(row=fila_inicio + 1, column=col_datos + 1, value=porcentaje_no_autorizadas)

    ws.cell(row=fila_inicio + 2, column=col_datos, value="OBSERVACIÓN")
    ws.cell(row=fila_inicio + 2, column=col_datos + 1, value=porcentaje_observacion)

    # Crear gráfico de pastel
    chart = PieChart()
    chart.title = titulo

    # Referencias a datos y etiquetas
    labels = Reference(ws, min_col=col_datos, min_row=fila_inicio, max_row=fila_inicio + 2)
    data = Reference(ws, min_col=col_datos + 1, min_row=fila_inicio, max_row=fila_inicio + 2)

    # Añadir datos y categorías correctamente
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(labels)

    # Configurar etiquetas de datos
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCategoryName = True
    chart.dataLabels.showVal = False
    chart.dataLabels.showSerName = False

    # Añadir gráfico a la hoja
    ws.add_chart(chart, celda_destino)

    return chart

# Agregar esta nueva función después de la función crear_grafico_pastel_openpyxl existente

def crear_grafico_pastel_estatus_openpyxl(ws, porcentaje_autorizadas, porcentaje_no_autorizadas, porcentaje_observacion, 
                                        titulo, celda_destino, identificador_unico):
    """
    Crea un gráfico de pastel para los estatus (autorizadas, no autorizadas, observación)
    """
    # Crear datos para el gráfico en celdas ocultas con posición única
    fila_inicio = 150 + (identificador_unico * 10)  # Espacio suficiente entre gráficos (diferente del anterior)
    col_datos = 25  # Columna Y para datos temporales

    # Limpiar celdas previas (por si acaso)
    for i in range(4):  # Limpiar 4 filas
        ws.cell(row=fila_inicio + i, column=col_datos, value="")
        ws.cell(row=fila_inicio + i, column=col_datos + 1, value="")

    # Crear datos para el gráfico
    ws.cell(row=fila_inicio, column=col_datos, value="AUTORIZADAS")
    ws.cell(row=fila_inicio, column=col_datos + 1, value=porcentaje_autorizadas)

    ws.cell(row=fila_inicio + 1, column=col_datos, value="NO AUTORIZADAS")
    ws.cell(row=fila_inicio + 1, column=col_datos + 1, value=porcentaje_no_autorizadas)

    ws.cell(row=fila_inicio + 2, column=col_datos, value="OBSERVACIÓN")
    ws.cell(row=fila_inicio + 2, column=col_datos + 1, value=porcentaje_observacion)

    # Crear gráfico de pastel
    chart = PieChart()
    chart.title = titulo

    # Referencias a datos y etiquetas
    labels = Reference(ws, min_col=col_datos, min_row=fila_inicio, max_row=fila_inicio + 2)
    data = Reference(ws, min_col=col_datos + 1, min_row=fila_inicio, max_row=fila_inicio + 2)

    # Añadir datos y categorías correctamente
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(labels)

    # Configurar etiquetas de datos
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCategoryName = True
    chart.dataLabels.showVal = False
    chart.dataLabels.showSerName = False

    # Añadir gráfico a la hoja
    ws.add_chart(chart, celda_destino)

    return chart


def cargar_configuracion():
    """Cargar configuración desde archivo JSON"""
    config_default = {
        "fm_path": "MedicionesFmCSV",
        "tv_path": "MedicionesTvCSV",
        "am_path": "MedicionesAmCSV",  # NUEVO
        "output_path": "ReportesOcupacion",
        "emisoras_por_ciudad": {}
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
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
ruta_am = config.get("am_path", "MedicionesAmCSV")
ruta_salida = config.get("output_path", "ReportesOcupacion")
fecha_actual = datetime.now().strftime("%d/%m/%Y")
# Umbral para AM (puedes ajustarlo según necesidades)
UMBRAL_AM = 45.0
# ------------------ FUNCIONES AUXILIARES ------------------

def crear_tabla_ocupacion_am(ws, datos, umbral=45, ciudad=""):
    """
    Crea la tabla de ocupación AM a partir de la columna J
    Similar a la función de FM
    """
    # Obtener los datos de la hoja
    fila_inicio = 2
    
    # Calcular estadísticas
    total_frecuencias = len(datos)
    
    # Lista para frecuencias problemáticas
    frecuencias_problematicas = []
    
    # Contar frecuencias operando mayor al umbral (OCUPACIÓN > 0%)
    frecuencias_mayor_umbral = 0
    for fila in range(fila_inicio, ws.max_row + 1):
        ocupacion_celda = ws.cell(row=fila, column=4)  # Columna D = Ocupación (%)
        if ocupacion_celda.value and isinstance(ocupacion_celda.value, (int, float)):
            if ocupacion_celda.value > 0:  # Ocupación > 0%
                frecuencias_mayor_umbral += 1
    
    # Contar frecuencias con criterios corregidos y detectar problemáticas
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
        
        # Verificar si tiene nombre en ESTACIÓN y si termina en "_OBSERVACION"
        tiene_nombre = estacion_celda.value and estacion_celda.value != "No identificada" and estacion_celda.value != ""
        es_observacion = False
        if tiene_nombre:
            estacion_str = str(estacion_celda.value).strip()
            es_observacion = estacion_str.upper().endswith("_OBSERVACION")
        
        # LÓGICA DE CLASIFICACIÓN (igual que FM)
        
        # Caso 1: Frecuencias con ocupación > 0% y nombre no termina en "_OBSERVACION"
        if ocupacion_valor > 0 and tiene_nombre and not es_observacion:
            estacion_str = str(estacion_celda.value).lower()
            if "no autorizado" in estacion_str or "no autorizada" in estacion_str or "no aut" in estacion_str:
                frecuencias_no_autorizadas += 1
                ws.cell(row=fila, column=1).fill = ROJO
                ws.cell(row=fila, column=2).fill = ROJO
                ws.cell(row=fila, column=4).fill = ROJO
            else:
                frecuencias_autorizadas += 1
                ws.cell(row=fila, column=1).fill = VERDE
                ws.cell(row=fila, column=2).fill = VERDE
                ws.cell(row=fila, column=4).fill = VERDE
        
        # Caso 2: Frecuencias con ocupación 0% y nombre no termina en "_OBSERVACION"
        elif ocupacion_valor == 0 and tiene_nombre and not es_observacion:
            estacion_str = str(estacion_celda.value).lower()
            if "no autorizado" in estacion_str or "no autorizada" in estacion_str or "no aut" in estacion_str:
                frecuencias_no_autorizadas += 1
                ws.cell(row=fila, column=1).fill = ROJO
                ws.cell(row=fila, column=2).fill = ROJO
                ws.cell(row=fila, column=4).fill = ROJO
            else:
                frecuencias_autorizadas += 1
                ws.cell(row=fila, column=1).fill = VERDE
                ws.cell(row=fila, column=2).fill = VERDE
                ws.cell(row=fila, column=4).fill = VERDE
            
            # DETECTAR COMO PROBLEMÁTICA
            frecuencia_celda = ws.cell(row=fila, column=1)
            level_celda = ws.cell(row=fila, column=5)
            
            frecuencia_problematica = {
                "ciudad": normalizar_nombre_ciudad(ciudad),
                "estacion": str(estacion_celda.value),
                "tipo": "AM",
                "frecuencia": frecuencia_celda.value if frecuencia_celda.value else "N/A",
                "ocupacion": 0,
                "level": level_celda.value if level_celda.value else "N/A",
                "fila_excel": fila
            }
            frecuencias_problematicas.append(frecuencia_problematica)
        
        # Caso 3: Frecuencias con ocupación > 0% sin nombre o con nombre terminado en "_OBSERVACION"
        elif ocupacion_valor > 0 and (not tiene_nombre or es_observacion):
            frecuencias_observacion += 1
            ws.cell(row=fila, column=1).fill = AMARILLO
            ws.cell(row=fila, column=2).fill = AMARILLO
            ws.cell(row=fila, column=4).fill = AMARILLO
        
        # Caso 4: Frecuencias con ocupación 0% sin nombre
        if ocupacion_valor == 0:
            frecuencias_libres += 1
    
    # Calcular porcentajes
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
    alignment_center = Alignment(horizontal="center", vertical="center")
    alignment_left = Alignment(horizontal="left", vertical="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Título
    ws.cell(row=fila_inicio_tabla, column=col_inicio, value="OCUPACIÓN AM")
    ws.merge_cells(start_row=fila_inicio_tabla, start_column=col_inicio, end_row=fila_inicio_tabla, end_column=col_inicio + 3)
    titulo_cell = ws.cell(row=fila_inicio_tabla, column=col_inicio)
    titulo_cell.font = Font(bold=True, size=14)
    titulo_cell.alignment = alignment_center
    titulo_cell.fill = GRIS
    
    # Encabezados de la tabla
    datos_tabla = [
        ["UMBRAL", umbral, "% FRECUENCIAS OCUPADAS", f"{porcentaje_ocupadas:.2f}%"],
        ["TOTAL DE FRECUENCIAS MONITOREADAS", total_frecuencias, "% FRECUENCIAS LIBRES", f"{porcentaje_libres:.2f}%"],
        ["FRECUENCIAS OCUPADAS", frecuencias_mayor_umbral, "% AUTORIZADAS", f"{porcentaje_autorizadas:.2f}%"],
        ["FRECUENCIAS AUTORIZADAS", frecuencias_autorizadas, "% NO AUTORIZADAS", f"{porcentaje_no_autorizadas:.2f}%"],
        ["FRECUENCIAS EN OBSERVACIÓN", frecuencias_observacion, "% INTERMODULACIÓN O RUIDO", f"{porcentaje_observacion:.2f}%"],
        ["FRECUENCIAS NO AUTORIZADOS", frecuencias_no_autorizadas, "", ""],
        ["FRECUENCIAS LIBRES", frecuencias_libres, "", ""]
    ]
    
    # Escribir datos de la tabla (similar a FM)
    for i, fila_datos in enumerate(datos_tabla, start=fila_inicio_tabla + 1):
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
    
    # Crear gráfico de pastel para AM
    try:
        crear_grafico_pastel_openpyxl(ws, porcentaje_ocupadas, porcentaje_libres, "OCUPACIÓN AM", "P2", 997)
    except Exception as e:
        print(f"Error al crear gráfico AM: {e}")

    # Crear gráfico de pastel para estatus AM
    try:
        crear_grafico_pastel_estatus_openpyxl(ws, porcentaje_autorizadas, porcentaje_no_autorizadas, 
                                            porcentaje_observacion, "ESTATUS AM", "Y2", 996)
    except Exception as e:
        print(f"Error al crear gráfico AM (estatus): {e}")

    return {
        "total_frecuencias": total_frecuencias,
        "frecuencias_mayor_umbral": frecuencias_mayor_umbral,
        "frecuencias_autorizadas": frecuencias_autorizadas,
        "frecuencias_no_autorizadas": frecuencias_no_autorizadas,
        "frecuencias_observacion": frecuencias_observacion,
        "frecuencias_libres": frecuencias_libres,
        "frecuencias_problematicas": frecuencias_problematicas,
        "porcentajes": {
            "ocupadas": porcentaje_ocupadas,
            "libres": porcentaje_libres,
            "autorizadas": porcentaje_autorizadas,
            "no_autorizadas": porcentaje_no_autorizadas,
            "observacion": porcentaje_observacion
        }
    }


def crear_tabla_ocupacion_fm(ws, datos, umbral=60, ciudad=""):
    """
    Crea la tabla de ocupación FM a partir de la columna J
    Retorna: estadísticas y frecuencias problemáticas
    """
    # Obtener los datos de la hoja
    fila_inicio = 2
    
    # Calcular estadísticas
    total_frecuencias = len(datos)
    
    # Lista para frecuencias problemáticas
    frecuencias_problematicas = []
    
    # Contar frecuencias operando mayor al umbral (OCUPACIÓN > 0%)
    frecuencias_mayor_umbral = 0
    for fila in range(fila_inicio, ws.max_row + 1):
        ocupacion_celda = ws.cell(row=fila, column=4)  # Columna D = Ocupación (%)
        if ocupacion_celda.value and isinstance(ocupacion_celda.value, (int, float)):
            if ocupacion_celda.value > 0:  # Ocupación > 0%
                frecuencias_mayor_umbral += 1
    
    # Contar frecuencias con criterios corregidos y detectar problemáticas
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
        
        # Verificar si tiene nombre en ESTACIÓN y si termina en "_OBSERVACION"
        tiene_nombre = estacion_celda.value and estacion_celda.value != "No identificada" and estacion_celda.value != ""
        es_observacion = False
        if tiene_nombre:
            estacion_str = str(estacion_celda.value).strip()
            es_observacion = estacion_str.upper().endswith("_OBSERVACION")
        
        # NUEVA LÓGICA DE CLASIFICACIÓN
        
        # Caso 1: Frecuencias con ocupación > 0% y nombre no termina en "_OBSERVACION"
        if ocupacion_valor > 0 and tiene_nombre and not es_observacion:
            # Se clasifican como autorizadas/no autorizadas
            estacion_str = str(estacion_celda.value).lower()
            if "no autorizado" in estacion_str or "no autorizada" in estacion_str or "no aut" in estacion_str:
                frecuencias_no_autorizadas += 1
                # Pintar de rojo
                ws.cell(row=fila, column=1).fill = ROJO  # Frecuencia
                ws.cell(row=fila, column=2).fill = ROJO  # Estación
                ws.cell(row=fila, column=4).fill = ROJO  # Ocupación
            else:
                frecuencias_autorizadas += 1
                # Pintar de verde
                ws.cell(row=fila, column=1).fill = VERDE  # Frecuencia
                ws.cell(row=fila, column=2).fill = VERDE  # Estación
                ws.cell(row=fila, column=4).fill = VERDE  # Ocupación
        
        # Caso 2: Frecuencias con ocupación 0% y nombre no termina en "_OBSERVACION"
        elif ocupacion_valor == 0 and tiene_nombre and not es_observacion:
            # Se detectan como problemáticas pero se incluyen en autorizadas/no autorizadas
            estacion_str = str(estacion_celda.value).lower()
            if "no autorizado" in estacion_str or "no autorizada" in estacion_str or "no aut" in estacion_str:
                frecuencias_no_autorizadas += 1
                # Pintar de rojo (aunque tenga 0% de ocupación)
                ws.cell(row=fila, column=1).fill = ROJO  # Frecuencia
                ws.cell(row=fila, column=2).fill = ROJO  # Estación
                ws.cell(row=fila, column=4).fill = ROJO  # Ocupación
            else:
                frecuencias_autorizadas += 1
                # Pintar de verde (aunque tenga 0% de ocupación)
                ws.cell(row=fila, column=1).fill = VERDE  # Frecuencia
                ws.cell(row=fila, column=2).fill = VERDE  # Estación
                ws.cell(row=fila, column=4).fill = VERDE  # Ocupación
            
            # DETECTAR COMO PROBLEMÁTICA (ocupación 0% con nombre válido)
            frecuencia_celda = ws.cell(row=fila, column=1)  # Columna A = Frecuencia
            level_celda = ws.cell(row=fila, column=5)       # Columna E = Level
            
            frecuencia_problematica = {
                "ciudad": normalizar_nombre_ciudad(ciudad),
                "estacion": str(estacion_celda.value),
                "tipo": "FM",
                "frecuencia": frecuencia_celda.value if frecuencia_celda.value else "N/A",
                "ocupacion": 0,
                "level": level_celda.value if level_celda.value else "N/A",
                "fila_excel": fila
            }
            frecuencias_problematicas.append(frecuencia_problematica)
        
        # Caso 3: Frecuencias con ocupación > 0% sin nombre o con nombre terminado en "_OBSERVACION"
        elif ocupacion_valor > 0 and (not tiene_nombre or es_observacion):
            # Se clasifican como observación
            frecuencias_observacion += 1
            # Pintar de amarillo
            ws.cell(row=fila, column=1).fill = AMARILLO  # Frecuencia
            ws.cell(row=fila, column=2).fill = AMARILLO  # Estación
            ws.cell(row=fila, column=4).fill = AMARILLO  # Ocupación
        
        # Caso 4: Frecuencias con ocupación 0% sin nombre
        if ocupacion_valor == 0:  # and not tiene_nombre: 
            # Se clasifican como libres
            frecuencias_libres += 1
            # No se pinta (queda con formato por defecto)
    
    # VERIFICACIÓN: La suma debe coincidir con el total
    suma_categorias = (frecuencias_autorizadas + frecuencias_no_autorizadas + 
                      frecuencias_observacion + frecuencias_libres)
    #print(f"✅ Frecuencias totales: {suma_categorias}, Autorizadas: {frecuencias_autorizadas}, No autorizadas: {frecuencias_no_autorizadas}, Observación: {frecuencias_observacion}, Libres: {frecuencias_libres}")

    # El resto de la función permanece igual...
    # Calcular porcentajes
    porcentaje_ocupadas = (frecuencias_mayor_umbral / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_libres = (frecuencias_libres / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_autorizadas = (frecuencias_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_no_autorizadas = (frecuencias_no_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
    porcentaje_observacion = (frecuencias_observacion / total_frecuencias * 100) if total_frecuencias > 0 else 0

    # ... (el resto del código de la función permanece igual)

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
        ["FRECUENCIAS OCUPADAS", frecuencias_mayor_umbral, "% AUTORIZADAS", f"{porcentaje_autorizadas:.2f}%"],
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
    
     # Crear gráfico de pastel para FM usando openpyxl
    # En crear_tabla_ocupacion_fm, modificar la llamada al gráfico:
    try:
        crear_grafico_pastel_openpyxl(ws, porcentaje_ocupadas, porcentaje_libres, "OCUPACIÓN FM", "P2", 999)  # ID único para FM
    except Exception as e:
        print(f"Error al crear gráfico FM: {e}")

     # NUEVO: Crear gráfico de pastel para estatus (autorizadas/no autorizadas/observación)
    try:
        # Obtener porcentajes de las celdas M4, M5, M6
        porcentaje_autorizadas = porcentaje_autorizadas  # Ya calculado
        porcentaje_no_autorizadas = porcentaje_no_autorizadas  # Ya calculado  
        porcentaje_observacion = porcentaje_observacion  # Ya calculado
        
        # Crear el nuevo gráfico en columna Y (celda Y2)
        crear_grafico_pastel_estatus_openpyxl(ws, porcentaje_autorizadas, porcentaje_no_autorizadas, 
                                            porcentaje_observacion, "ESTATUS FM", "Y2", 998)
    except Exception as e:
        print(f"Error al crear gráfico FM (estatus): {e}")

    return {
        "total_frecuencias": total_frecuencias,
        "frecuencias_mayor_umbral": frecuencias_mayor_umbral,
        "frecuencias_autorizadas": frecuencias_autorizadas,
        "frecuencias_no_autorizadas": frecuencias_no_autorizadas,
        "frecuencias_observacion": frecuencias_observacion,
        "frecuencias_libres": frecuencias_libres,
        "frecuencias_problematicas": frecuencias_problematicas,  # ← NUEVO
        "porcentajes": {
            "ocupadas": porcentaje_ocupadas,
            "libres": porcentaje_libres,
            "autorizadas": porcentaje_autorizadas,
            "no_autorizadas": porcentaje_no_autorizadas,
            "observacion": porcentaje_observacion
        }
    }

def verificar_posicion_tablas(ws):
    """Verifica qué hay en las posiciones donde deberían estar las tablas"""
    print("=== VERIFICACIÓN DE POSICIONES ===")
    
    # Posiciones donde deberían estar las tablas
    posiciones_verificar = [
        (1, 11),  # K1 - Primera tabla
        (1, 16),  # P1 - Segunda tabla  
        (12, 11)  # K12 - Tercera tabla
    ]
    
    for fila, col in posiciones_verificar:
        celda = ws.cell(row=fila, column=col)
        print(f"Celda {get_column_letter(col)}{fila}: '{celda.value}'")
        

def crear_tablas_ocupacion_tv(ws, datos, ciudad=""):
    """
    Crea las tablas de ocupación TV por bandas a partir de la columna K
    """
    
    # Al inicio de la función
    #verificar_posicion_tablas(ws)
    
    # ... resto del código ...
    #debug_bandas_tv(ws)
    
    # Definir umbrales por banda
    umbrales_por_banda = {
        "Bandas I-III (VHF)": UMBRAL_TV_BANDA_I_III,
        "Banda III (VHF)": UMBRAL_TV_BANDA_III,
        "Bandas IV-V (UHF)": UMBRAL_TV_BANDA_IV_V
    }
    
  
    # Obtener los datos de la hoja
    fila_inicio = 2
    
    # Separar datos por banda (con manejo de variaciones)
    datos_por_banda = {banda: [] for banda in umbrales_por_banda.keys()}
    
    # Separar datos por banda
    datos_por_banda = {}
    for banda in umbrales_por_banda.keys():
        datos_por_banda[banda] = []
    
    for fila in range(fila_inicio, ws.max_row + 1):
        banda_celda = ws.cell(row=fila, column=3)  # Columna C = Banda
        if banda_celda.value in umbrales_por_banda:
            datos_por_banda[banda_celda.value].append(fila)
        
    
    # Crear tabla para cada banda
    col_inicio = 11  # Columna K
    fila_actual = 1
    separacion_entre_tablas = 2
    
    # Ordenar bandas para consistencia
    bandas_orden = ["Bandas I-III (VHF)", "Banda III (VHF)", "Bandas IV-V (UHF)"]
    
    resultados_bandas = {}

    # Filas fijas para cada gráfico
    filas_graficos = {
        "Bandas I-III (VHF)": 1,    # Primera gráfica en fila 1
        "Banda III (VHF)": 16,      # Segunda gráfica en fila 16  
        "Bandas IV-V (UHF)": 31     # Tercera gráfica en fila 31
    }

    # NUEVO: Mapeo de filas para los gráficos de estatus (en columna Y)
    filas_graficos_estatus = {
        "Bandas I-III (VHF)": 1,    # Misma fila que el gráfico original
        "Banda III (VHF)": 16,      # Misma fila que el gráfico original
        "Bandas IV-V (UHF)": 31     # Misma fila que el gráfico original
    }

    
    
    resultados_bandas = {}
    frecuencias_problematicas_tv = []  # ← NUEVO: lista para todas las frecuencias problemáticas de TV

    for banda, filas_banda in datos_por_banda.items():
        if not filas_banda:
            continue  # Saltar bandas sin datos
        
        umbral = umbrales_por_banda[banda]
        
        # Calcular estadísticas para esta banda
        total_frecuencias = len(filas_banda)
        
        frecuencias_mayor_umbral = 0
        frecuencias_problematicas_banda = []  # ← NUEVO: para esta banda específica
        
        for fila in filas_banda:
            ocupacion_celda = ws.cell(row=fila, column=5)  # Columna E = Ocupación (%) para TV
            if ocupacion_celda.value and isinstance(ocupacion_celda.value, (int, float)):
                if ocupacion_celda.value > 0:  # Ocupación > 0%
                    frecuencias_mayor_umbral += 1
        
        # Contar frecuencias con criterios específicos y detectar problemáticas
        frecuencias_autorizadas = 0
        frecuencias_no_autorizadas = 0
        frecuencias_observacion = 0
        frecuencias_libres = 0
        
        # En la función crear_tablas_ocupacion_tv, modifica la sección de clasificación:

        # ... código anterior ...

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
            
            # Verificar si tiene nombre en ESTACIÓN y si termina en "_OBSERVACION"
            tiene_nombre = estacion_celda.value and estacion_celda.value != "No identificada" and estacion_celda.value != ""
            es_observacion = False
            if tiene_nombre:
                estacion_str = str(estacion_celda.value).strip()
                es_observacion = estacion_str.upper().endswith("_OBSERVACION")
            
            # NUEVA LÓGICA DE CLASIFICACIÓN (IGUAL QUE PARA FM)
            
            # Caso 1: Frecuencias con ocupación > 0% y nombre no termina en "_OBSERVACION"
            if ocupacion_valor > 0 and tiene_nombre and not es_observacion:
                # Se clasifican como autorizadas/no autorizadas
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
            
            # Caso 2: Frecuencias con ocupación 0% y nombre no termina en "_OBSERVACION"
            elif ocupacion_valor == 0 and tiene_nombre and not es_observacion:
                # Se detectan como problemáticas pero se incluyen en autorizadas/no autorizadas
                estacion_str = str(estacion_celda.value).lower()
                if any(x in estacion_str for x in ["no autorizado", "no autorizada", "no aut"]):
                    frecuencias_no_autorizadas += 1
                    # Pintar de rojo (aunque tenga 0% de ocupación)
                    ws.cell(row=fila, column=1).fill = ROJO  # Frecuencia
                    ws.cell(row=fila, column=2).fill = ROJO  # Estación
                    ws.cell(row=fila, column=5).fill = ROJO  # Ocupación
                else:
                    frecuencias_autorizadas += 1
                    # Pintar de verde (aunque tenga 0% de ocupación)
                    ws.cell(row=fila, column=1).fill = VERDE  # Frecuencia
                    ws.cell(row=fila, column=2).fill = VERDE  # Estación
                    ws.cell(row=fila, column=5).fill = VERDE  # Ocupación
                
                # DETECTAR COMO PROBLEMÁTICA (ocupación 0% con nombre válido)
                frecuencia_celda = ws.cell(row=fila, column=1)  # Columna A = Frecuencia
                level_celda = ws.cell(row=fila, column=6)       # Columna F = Level
                banda_celda = ws.cell(row=fila, column=3)       # Columna C = Banda
                
                frecuencia_problematica = {
                    "ciudad": normalizar_nombre_ciudad(ciudad),
                    "estacion": str(estacion_celda.value),
                    "tipo": f"TV-{banda_celda.value}" if banda_celda.value else "TV-Desconocida",
                    "frecuencia": frecuencia_celda.value if frecuencia_celda.value else "N/A",
                    "ocupacion": 0,
                    "level": level_celda.value if level_celda.value else "N/A",
                    "fila_excel": fila
                }
                frecuencias_problematicas_banda.append(frecuencia_problematica)
                frecuencias_problematicas_tv.append(frecuencia_problematica)
            
            # Caso 3: Frecuencias con ocupación > 0% sin nombre o con nombre terminado en "_OBSERVACION"
            elif ocupacion_valor > 0 and (not tiene_nombre or es_observacion):
                # Se clasifican como observación
                frecuencias_observacion += 1
                # Pintar de amarillo
                ws.cell(row=fila, column=1).fill = AMARILLO  # Frecuencia
                ws.cell(row=fila, column=2).fill = AMARILLO  # Estación
                ws.cell(row=fila, column=5).fill = AMARILLO  # Ocupación
            
            # Caso 4: Frecuencias con ocupación 0% sin nombre
            if ocupacion_valor == 0: # and not tiene_nombre:
                # Se clasifican como libres
                frecuencias_libres += 1
                # No se pinta (queda con formato por defecto)

        # ... código posterior ...
        
        # Calcular porcentajes
        porcentaje_ocupadas = (frecuencias_mayor_umbral / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_libres = (frecuencias_libres / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_autorizadas = (frecuencias_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_no_autorizadas = (frecuencias_no_autorizadas / total_frecuencias * 100) if total_frecuencias > 0 else 0
        porcentaje_observacion = (frecuencias_observacion / total_frecuencias * 100) if total_frecuencias > 0 else 0
        
        # Guardar resultados para esta banda
        resultados_bandas[banda] = {
            "total_frecuencias": total_frecuencias,
            "frecuencias_mayor_umbral": frecuencias_mayor_umbral,
            "frecuencias_autorizadas": frecuencias_autorizadas,
            "frecuencias_no_autorizadas": frecuencias_no_autorizadas,
            "frecuencias_observacion": frecuencias_observacion,
            "frecuencias_libres": frecuencias_libres,
            "frecuencias_problematicas": frecuencias_problematicas_banda,  # ← NUEVO
            "porcentajes": {
                "ocupadas": porcentaje_ocupadas,
                "libres": porcentaje_libres,
                "autorizadas": porcentaje_autorizadas,
                "no_autorizadas": porcentaje_no_autorizadas,
                "observacion": porcentaje_observacion
            }
        }
        
        # ... (resto de la función igual, creando la tabla) ...
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
        titulo_cell.font = Font(bold=True, size=12)
        titulo_cell.alignment = alignment_center
        titulo_cell.fill = GRIS
        
        # Encabezados de la tabla
        datos_tabla = [
            ["UMBRAL", umbral, "% FRECUENCIAS OCUPADAS", f"{porcentaje_ocupadas:.2f}%"],
            ["TOTAL DE FRECUENCIAS MONITOREADAS", total_frecuencias, "% FRECUENCIAS LIBRES", f"{porcentaje_libres:.2f}%"],
            ["FRECUENCIAS OCUPADAS", frecuencias_mayor_umbral, "% AUTORIZADAS", f"{porcentaje_autorizadas:.2f}%"],
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
        
        # Crear gráfico de pastel para esta banda de TV usando openpyxl
        try:
            # Usar fila fija según la banda
            fila_grafico = filas_graficos.get(banda, fila_actual)
            celda_destino = f"P{fila_grafico}"
            
            # Usar un identificador único para cada gráfico basado en el nombre de la banda
            identificador_unico = hash(banda) % 100  # Hash único para cada banda
            
            crear_grafico_pastel_openpyxl(ws, porcentaje_ocupadas, porcentaje_libres, 
                                        f"OCUPACIÓN {banda.upper()}", celda_destino, identificador_unico)
        except Exception as e:
            print(f"Error al crear gráfico para {banda}: {e}")
        
         # NUEVO: Crear gráfico de pastel para estatus de esta banda
        try:
            fila_grafico_estatus = filas_graficos_estatus.get(banda, fila_actual)
            celda_destino_estatus = f"Y{fila_grafico_estatus}"
            
            # Usar un identificador diferente para el gráfico de estatus
            identificador_unico_estatus = (hash(banda) + 1000) % 100  # Diferente del anterior
            
            # Determinar de qué celdas obtener los porcentajes según la banda
            if banda == "Bandas I-III (VHF)":
                # Usar porcentajes calculados: N4, N5, N6
                porcentaje_autorizadas_banda = porcentaje_autorizadas
                porcentaje_no_autorizadas_banda = porcentaje_no_autorizadas
                porcentaje_observacion_banda = porcentaje_observacion
            elif banda == "Banda III (VHF)":
                # Usar porcentajes calculados: N14, N15, N16  
                porcentaje_autorizadas_banda = porcentaje_autorizadas
                porcentaje_no_autorizadas_banda = porcentaje_no_autorizadas
                porcentaje_observacion_banda = porcentaje_observacion
            elif banda == "Bandas IV-V (UHF)":
                # Usar porcentajes calculados: N24, N25, N26
                porcentaje_autorizadas_banda = porcentaje_autorizadas
                porcentaje_no_autorizadas_banda = porcentaje_no_autorizadas
                porcentaje_observacion_banda = porcentaje_observacion
            
            crear_grafico_pastel_estatus_openpyxl(ws, porcentaje_autorizadas_banda, porcentaje_no_autorizadas_banda,
                                                porcentaje_observacion_banda, f"ESTATUS {banda.upper()}", 
                                                celda_destino_estatus, identificador_unico_estatus)
        except Exception as e:
            print(f"Error al crear gráfico de estatus para {banda}: {e}")




        # Actualizar fila actual para la próxima tabla
        fila_actual = fila_actual + len(datos_tabla) + separacion_entre_tablas + 1

        #print(f"✅ Tabla creada para {banda}: {total_frecuencias} frecuencias")
    
    return resultados_bandas, frecuencias_problematicas_tv  # ← MODIFICADO


def buscar_emisora_por_frecuencia(ciudad, frecuencia, tipo, tolerancia=0.1):
    """
    Busca una emisora por frecuencia en una ciudad específica
    tolerancia: margen de error en MHz para coincidir frecuencias
    """
    config = cargar_configuracion()
    emisoras_por_ciudad = config.get("emisoras_por_ciudad", {})
    
    # Normalizar nombre de ciudad
    ciudad_normalizada = ciudad.lower().strip()
    
    # Buscar por similitud si no se encuentra exacto
    if ciudad_normalizada not in emisoras_por_ciudad:
        for clave_real in emisoras_por_ciudad.keys():
            if ciudad.lower() in clave_real.lower() or clave_real.lower() in ciudad.lower():
                ciudad_normalizada = clave_real
                break
        else:
            return None
    
    emisoras = emisoras_por_ciudad[ciudad_normalizada].get(tipo, [])
    
    if not emisoras:
        return None
    
    # CASO ESPECÍFICO PARA AM - TOLERANCIAS MÁS ESTRICTAS
    if tipo == "AM":
        # Para AM: coincidencia exacta muy estricta primero
        for emisora in emisoras:
            try:
                freq_emisora = float(emisora.get("frecuencia", 0))
                if abs(freq_emisora - frecuencia) <= 0.001:  # Coincidencia muy exacta (0.001 MHz)
                    return emisora.get("nombre", "Desconocido")
            except (ValueError, TypeError):
                continue
        
        # Luego buscar con tolerancia muy reducida
        for emisora in emisoras:
            try:
                freq_emisora = float(emisora.get("frecuencia", 0))
                if abs(freq_emisora - frecuencia) <= 0.005:  # Tolerancia muy reducida (0.01 MHz)
                    return emisora.get("nombre", "Desconocido")
            except (ValueError, TypeError):
                continue
        
        return None
    
    else:
        # PARA FM Y TV - MANTENER TOLERANCIAS ORIGINALES
        # Buscar coincidencia exacta primero
        for emisora in emisoras:
            try:
                freq_emisora = float(emisora.get("frecuencia", 0))
                if abs(freq_emisora - frecuencia) <= 0.01:  # Coincidencia exacta
                    return emisora.get("nombre", "Desconocido")
            except (ValueError, TypeError):
                continue
        
        # Luego buscar con tolerancia normal
        for emisora in emisoras:
            try:
                freq_emisora = float(emisora.get("frecuencia", 0))
                if abs(freq_emisora - frecuencia) <= tolerancia:  # Tolerancia normal (0.1 MHz)
                    return emisora.get("nombre", "Desconocido")
            except (ValueError, TypeError):
                continue
    
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
    os.makedirs(ruta_am, exist_ok=True)  # NUEVO

def obtener_codigo_base(base):
    """Obtiene el código correspondiente según el nombre de la base"""
    correspondencia = {
        "zamora": "SCS-L01",
        "loja": "SCS-L02", 
        #"cañar": "SCS-L03",  # ñ normal
        #"cañar": "SCS-L03",  # ñ con tilde combinable (n + ˜)
        "tambo":"SCS-L03",
        "macas": "SCS-L04",
        "machala": "SCC-L04",
        "cuenca": "SCS-L05"
    }
    
    # Normalizar el nombre de la base
    base_normalizada = base.lower().strip()
    
    # Manejar diferentes representaciones de "cañar"
    """if (base_normalizada == "cañar" or 
        base_normalizada == "cañar" or  # ñ con tilde combinable
        base_normalizada == "canar" or   # sin tilde
        base_normalizada == "caÃ±ar"):   # posible encoding issue
        base_normalizada = "cañar"
    """
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


def reducir_archivo_am_csv(ruta_archivo):
    """
    Reduce el tamaño del archivo CSV AM conservando hasta la frecuencia 1560000 Hz (fila 104)
    """
    try:
        # Leer el archivo completo
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Convertir frecuencia a MHz
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        
        # Filtrar hasta 1560000 Hz (1.56 MHz) o fila 104
        df_filtrado = df[df["Frecuencia (MHz)"] <= 1.56]
        if len(df_filtrado) > 104:
            df_filtrado = df_filtrado.head(104)
        
        # Eliminar la columna temporal
        if "Frecuencia (MHz)" in df_filtrado.columns:
            df_filtrado = df_filtrado.drop(columns=["Frecuencia (MHz)"])
        
        # Guardar el archivo reducido (sobreescribir el original)
        df_filtrado.to_csv(ruta_archivo, index=False, encoding='latin-1')
        
        return True
        
    except Exception as e:
        print(f"Error reduciendo archivo AM {ruta_archivo}: {e}")
        return False
    




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
            df_filtrado = df[(df["Frecuencia (MHz)"] >= 88.1) & (df["Frecuencia (MHz)"] <= 108)]
            if len(df_filtrado) > 101:
                df_filtrado = df_filtrado.head(101)
        else:  # TV
            # Para TV: rango 55.25-693.25 MHz, conservar primeras 45 filas
            df_filtrado = df[(df["Frecuencia (MHz)"] >= 55.25) & (df["Frecuencia (MHz)"] <= 693.25)]
            if len(df_filtrado) > 46:
                df_filtrado = df_filtrado.head(46)
        
        # Eliminar la columna temporal
        if "Frecuencia (MHz)" in df_filtrado.columns:
            df_filtrado = df_filtrado.drop(columns=["Frecuencia (MHz)"])
        
        # Guardar el archivo reducido (sobreescribir el original)
        df_filtrado.to_csv(ruta_archivo, index=False, encoding='latin-1')
        
        return True
        
    except Exception as e:
        print(f"Error reduciendo archivo {ruta_archivo}: {e}")
        return False


def formatear_hoja_ocupacion_am(ws, datos, ciudad):
    """Formatea una hoja de ocupación AM con bordes y estilos"""
    # Limpiar hoja existente
    ws.delete_rows(1, ws.max_row)
    
    # Agregar encabezado para AM
    encabezados = [
        "Frecuencia (MHz)", "Estación", "FECHA DE SUSCRIPCION", "Ocupación (%)",
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
        nombre_emisora = buscar_emisora_por_frecuencia(ciudad, frecuencia, "AM")
        
        ws.cell(row=fila_idx, column=1, value=frecuencia)
        ws.cell(row=fila_idx, column=2, value=nombre_emisora or "No identificada")
        ws.cell(row=fila_idx, column=3, value=fila["FECHA DE SUSCRIPCION"])
        ws.cell(row=fila_idx, column=4, value=fila["Ocupación (%)"])
        ws.cell(row=fila_idx, column=5, value=fila["Level (dBµV/m)"])
        ws.cell(row=fila_idx, column=6, value=fila["Bandwidth (Hz)"])
        ws.cell(row=fila_idx, column=7, value=fila["Offset (Hz)"])
        ws.cell(row=fila_idx, column=8, value=fila["AM (%)"])
    
    # Aplicar bordes y formato (similar a FM)
    thin = Side(border_style="thin")
    borde_grueso = Side(border_style="medium")
    
    num_columnas = len(encabezados)
    num_filas = len(datos) + 1
    
    for row in range(1, num_filas + 1):
        for col in range(1, num_columnas + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            if row == 1:
                cell.border = Border(top=borde_grueso, bottom=thin,
                                   left=borde_grueso if col == 1 else thin,
                                   right=borde_grueso if col == num_columnas else thin)
            elif row == num_filas:
                cell.border = Border(top=thin, bottom=borde_grueso,
                                   left=borde_grueso if col == 1 else thin,
                                   right=borde_grueso if col == num_columnas else thin)
            elif col == 1:
                cell.border = Border(left=borde_grueso, top=thin, bottom=thin, right=thin)
            elif col == num_columnas:
                cell.border = Border(right=borde_grueso, top=thin, bottom=thin, left=thin)
    
    # Ajustar anchos de columnas
    for col in range(1, num_columnas + 1):
        col_letter = get_column_letter(col)
        max_length = 0
        for cell in ws[col_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        if col == 2:  # Columna Estación
            ws.column_dimensions[col_letter].width = max(max_length + 2, 25)
        else:
            ws.column_dimensions[col_letter].width = max_length + 2
    
    # Crear tabla de ocupación AM
    resultados_am = crear_tabla_ocupacion_am(ws, datos, ciudad=ciudad)
    frecuencias_problematicas = resultados_am.get("frecuencias_problematicas", [])
    
    # Quitar sufijo _OBSERVACION después de la clasificación
    for fila in range(2, ws.max_row + 1):
        estacion_celda = ws.cell(row=fila, column=2)
        if estacion_celda.value and "_OBSERVACION" in estacion_celda.value:
            nombre_limpio = estacion_celda.value.replace("_OBSERVACION", "").strip()
            estacion_celda.value = nombre_limpio

    return frecuencias_problematicas




def formatear_hoja_ocupacion(ws, datos, tipo, ciudad):
    """Formatea una hoja de ocupación con bordes y estilos, incluyendo columna Estación
    Retorna: frecuencias problemáticas detectadas"""
    
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
        
        # Buscar emisora por frecuencia (SIN quitar _OBSERVACION para la clasificación)
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
    
    # MODIFICACIÓN: Ahora las funciones de creación de tablas retornan frecuencias problemáticas
    frecuencias_problematicas = []
    
    if tipo == "FM":
        resultados_fm = crear_tabla_ocupacion_fm(ws, datos, ciudad=ciudad)
        frecuencias_problematicas = resultados_fm.get("frecuencias_problematicas", [])
    elif tipo == "TV":
        resultados_tv, frecuencias_problematicas_tv = crear_tablas_ocupacion_tv(ws, datos, ciudad=ciudad)
        frecuencias_problematicas = frecuencias_problematicas_tv

    # NUEVO: QUITAR EL SUFIJO _OBSERVACION DESPUÉS DE LA CLASIFICACIÓN
    for fila in range(2, ws.max_row + 1):
        estacion_celda = ws.cell(row=fila, column=2)  # Columna B = Estación
        if estacion_celda.value and "_OBSERVACION" in estacion_celda.value:
            # Quitar el sufijo solo para visualización, después de que ya se hizo la clasificación
            nombre_limpio = estacion_celda.value.replace("_OBSERVACION", "").strip()
            estacion_celda.value = nombre_limpio

    return frecuencias_problematicas


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
    """Determina la banda según la frecuencia para TV"""
    if freq is None:
        return "Otra banda"
    
    try:
        freq = float(freq)
    except (ValueError, TypeError):
        return "Otra banda"
    
    # Bandas de TV según estándares
    if (54 <= freq <= 72) or (76 <= freq <= 88):
        return "Bandas I-III (VHF)"
    elif 174 <= freq <= 216:
        return "Banda III (VHF)"
    elif (470 <= freq <= 488) or (512 <= freq <= 608) or (614 <= freq <= 698):
        return "Bandas IV-V (UHF)"
    else:
        return "Otra banda"
    


def procesar_archivo_am(ruta_archivo, base):
    """Procesa archivo AM y extrae datos de ocupación hasta 1.56 MHz"""
    try:
        # Primero reducir el archivo
        if not reducir_archivo_am_csv(ruta_archivo):
            return None, None
            
        # Leer el archivo reducido
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Filtrar por rango de frecuencia (hasta 1.56 MHz)
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        df_filtrado = df[df["Frecuencia (MHz)"] <= 1.56]
        
        if df_filtrado.empty:
            return None, None
        
        # Obtener mes de los datos
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
        
        # Buscar las columnas necesarias (similar a FM)
        columna_ocupacion = buscar_columna_por_patron(df_filtrado, ['ocupaci'])
        columna_level = buscar_columna_por_patron(df_filtrado, ['level_3', 'nivel'])
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
                        level_val = limpiar_valor_numerico(row[columna_level]) if columna_level else np.nan
                        bandwidth_val = limpiar_valor_numerico(row[columna_bandwidth]) if columna_bandwidth else np.nan
                        offset_val = limpiar_valor_numerico(row[columna_offset]) if columna_offset else np.nan
                        am_val = limpiar_valor_numerico(row[columna_am]) if columna_am else np.nan
                        
                        ocupacion_data.append({
                            "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                            "FECHA DE SUSCRIPCION": datetime.now().strftime("%Y-%m-%d"),
                            "Ocupación (%)": ocupacion_val,
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
                    level_val = limpiar_valor_numerico(row[columna_level]) if columna_level else np.nan
                    bandwidth_val = limpiar_valor_numerico(row[columna_bandwidth]) if columna_bandwidth else np.nan
                    offset_val = limpiar_valor_numerico(row[columna_offset]) if columna_offset else np.nan
                    am_val = limpiar_valor_numerico(row[columna_am]) if columna_am else np.nan
                    
                    ocupacion_data.append({
                        "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                        "FECHA DE SUSCRIPCION": datetime.now().strftime("%Y-%m-%d"),
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
        print(f"Error procesando archivo AM {ruta_archivo}: {e}")
        import traceback
        traceback.print_exc()
        return None, base
    

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
        columna_level = buscar_columna_por_patron(df_filtrado, ['level_3', 'nivel'])
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
        columna_level = buscar_columna_por_patron(df_filtrado, ['level_3', 'nivel'])
        columna_bandwidth = buscar_columna_por_patron(df_filtrado, ['bandwidth', 'ancho de banda'])
        columna_offset = buscar_columna_por_patron(df_filtrado, ['offset', 'desplazamiento'])
        columna_am = buscar_columna_por_patron(df_filtrado, ['am', 'amplitud modulada'])
        
        # Calcular ocupación usando los valores reales del archivo
        ocupacion_data = []
        
        if columna_ocupacion:
            # Usar los valores reales de ocupación del archivo
            for _, row in df_filtrado.iterrows():
                frecuencia = row["Frecuencia (MHz)"]
                # Determinar la banda según la frecuencia si no está definida
                banda_val = obtener_banda_por_frecuencia(frecuencia)
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

def debug_bandas_tv(ws):
    """Debug: mostrar información sobre las bandas TV encontradas"""
    print("=== DEBUG: BANDAS TV ENCONTRADAS ===")
    
    # Contar frecuencias por banda
    bandas_count = {}
    fila_inicio = 2
    
    for fila in range(fila_inicio, ws.max_row + 1):
        banda_celda = ws.cell(row=fila, column=3)  # Columna C = Banda
        if banda_celda.value:
            banda = banda_celda.value
            bandas_count[banda] = bandas_count.get(banda, 0) + 1
    
    print("Frecuencias por banda:")
    for banda, count in bandas_count.items():
        print(f"  {banda}: {count} frecuencias")
    
    # Verificar también las frecuencias específicas
    print("\nFrecuencias en rango de Banda III (174-216 MHz):")
    for fila in range(fila_inicio, ws.max_row + 1):
        freq_celda = ws.cell(row=fila, column=1)  # Columna A = Frecuencia (MHz)
        banda_celda = ws.cell(row=fila, column=3)  # Columna C = Banda
        if freq_celda.value and isinstance(freq_celda.value, (int, float)):
            if 174 <= freq_celda.value <= 216:
                print(f"  Fila {fila}: {freq_celda.value} MHz - Banda: {banda_celda.value}")


def verificar_posicion_tablas(ws):
    """Verifica qué hay en las posiciones donde deberían estar las tablas"""
    print("=== VERIFICACIÓN DE POSICIONES ===")
    
    # Posiciones donde deberían estar las tablas
    posiciones_verificar = [
        (1, 11),  # K1 - Primera tabla
        (1, 16),  # P1 - Segunda tabla  
        (12, 11)  # K12 - Tercera tabla
    ]
    
    for fila, col in posiciones_verificar:
        celda = ws.cell(row=fila, column=col)
        print(f"Celda {get_column_letter(col)}{fila}: '{celda.value}'")












# ------------------ FUNCIÓN PRINCIPAL DE PROCESAMIENTO ------------------


# Modificar la función procesar_ocupacion para aceptar el parámetro umbrales
def procesar_ocupacion(callback_progreso=None, callback_log=None, umbrales=None):
    """
    Función principal que procesa datos de ocupación de espectro
    Retorna: dict con información de procesamiento y frecuencias problemáticas
    """
    # Inicializar directorios
    inicializar_directorios()
    if callback_log:
        callback_log("=== VERIFICACIÓN INICIAL DE ARCHIVOS ===")
    
    # Obtener listas de archivos
    try:
        archivos_fm = {obtener_base(f): os.path.join(ruta_fm, f) for f in os.listdir(ruta_fm) if f.endswith(".csv")}
        archivos_tv = {obtener_base(f): os.path.join(ruta_tv, f) for f in os.listdir(ruta_tv) if f.endswith(".csv")}
        archivos_am = {obtener_base(f): os.path.join(ruta_am, f) for f in os.listdir(ruta_am) if f.endswith(".csv")}
        
        if callback_log:
            callback_log(f"Archivos FM encontrados: {list(archivos_fm.keys())}")
            callback_log(f"Archivos TV encontrados: {list(archivos_tv.keys())}")
            callback_log(f"Archivos AM encontrados: {list(archivos_am.keys())}")
            
    except Exception as e:
        if callback_log:
            callback_log(f"Error al leer archivos: {str(e)}")
        return {"error": f"Error al leer archivos: {str(e)}"}
    

    # Inicializar estructura para resultados
    resultado = {
        "existen_problemas": False,
        "datos": {"FM": [], "TV": [], "AM": []},  # NUEVO: Agregar AM
        "archivos_generados": [],
        "error": None
    }
    
    if callback_log:
        callback_log("Iniciando procesamiento de ocupación de espectro...")
        callback_log(f"Ruta FM: {ruta_fm}")
        callback_log(f"Ruta TV: {ruta_tv}")
        callback_log(f"Ruta AM: {ruta_am}")  # NUEVO
        callback_log(f"Ruta salida: {ruta_salida}")
    
    # Si no se proporcionan umbrales, usar los globales por defecto
    if umbrales is None:
        umbrales = {
            "global": {
                "FM": 60.0,
                "AM": 45.0,  # NUEVO: Umbral para AM
                "TV": {
                    "tipo": "general",
                    "valor": 45.0,
                    "valores": {
                        "Banda I-III": 47.0,
                        "Banda III": 56.0,
                        "Banda IV-V": 64.0
                    }
                }
            }
        }
        if callback_log:
            callback_log("⚠️  Usando umbrales globales por defecto")
    
    # Obtener listas de archivos
    try:
        archivos_fm = {obtener_base(f): os.path.join(ruta_fm, f) for f in os.listdir(ruta_fm) if f.endswith(".csv")}
        archivos_tv = {obtener_base(f): os.path.join(ruta_tv, f) for f in os.listdir(ruta_tv) if f.endswith(".csv")}
        archivos_am = {obtener_base(f): os.path.join(ruta_am, f) for f in os.listdir(ruta_am) if f.endswith(".csv")}  # NUEVO
        
        if callback_log:
            callback_log(f"Encontrados {len(archivos_fm)} archivos FM, {len(archivos_tv)} archivos TV y {len(archivos_am)} archivos AM")  # MODIFICADO
    except Exception as e:
        if callback_log:
            callback_log(f"Error al leer archivos: {str(e)}")
        resultado["error"] = f"Error al leer archivos: {str(e)}"
        return resultado

    # Encontrar bases comunes (para FM y TV)
    bases_comunes = set(archivos_fm.keys()).intersection(archivos_tv.keys())
    total_bases = len(bases_comunes)
    
    if total_bases == 0:
        if callback_log:
            callback_log("No se encontraron bases comunes entre FM y TV")
        resultado["error"] = "No se encontraron bases comunes entre FM y TV"
        return resultado
    
    if callback_log:
        callback_log(f"Procesando {total_bases} bases comunes")
    
    # Variables para tracking
    archivos_generados = []
    todas_frecuencias_problematicas = {"FM": [], "TV": [], "AM": []}  # NUEVO: Agregar AM
    
    # Procesar cada base común
    # En la función procesar_ocupacion, dentro del loop de bases_comunes:
    for i, base in enumerate(bases_comunes):
        if callback_progreso:
            progreso = int((i / total_bases) * 100)
            callback_progreso(progreso)
            
        if callback_log:
            callback_log(f"Procesando base: {base}")
        
        try:
            # VERIFICACIÓN DIAGNÓSTICA
            if callback_log:
                callback_log(f"🔍 Verificando archivos para {base}:")
                callback_log(f"   FM: {'SÍ' if base in archivos_fm else 'NO'}")
                callback_log(f"   TV: {'SÍ' if base in archivos_tv else 'NO'}")
                callback_log(f"   AM: {'SÍ' if base in archivos_am else 'NO'}")
            
            # Procesar archivos FM y TV
            datos_fm, base_fm = procesar_archivo_fm(archivos_fm[base], base)
            datos_tv, base_tv = procesar_archivo_tv(archivos_tv[base], base)
            
            # NUEVO: Procesar AM solo para Cuenca
            datos_am = None
            if base.lower() == "cuenca" and base in archivos_am:
                if callback_log:
                    callback_log(f"📻 Procesando archivo AM para Cuenca: {archivos_am[base]}")
                datos_am, base_am = procesar_archivo_am(archivos_am[base], base)
                if datos_am is not None and not datos_am.empty:
                    if callback_log:
                        callback_log(f"✅ Datos AM procesados para Cuenca: {len(datos_am)} frecuencias")
                else:
                    if callback_log:
                        callback_log(f"⚠️  No se pudieron procesar datos AM para {base}")
                    datos_am = pd.DataFrame(columns=[
                        "Frecuencia (MHz)", "FECHA DE SUSCRIPCION", "Ocupación (%)",
                        "Level (dBµV/m)", "Bandwidth (Hz)", "Offset (Hz)", "AM (%)", "Mes"
                    ])
            else:
                datos_am = pd.DataFrame(columns=[
                    "Frecuencia (MHz)", "FECHA DE SUSCRIPCION", "Ocupación (%)",
                    "Level (dBµV/m)", "Bandwidth (Hz)", "Offset (Hz)", "AM (%)", "Mes"
    ])
            


            # DIAGNÓSTICO DE DATOS
            if callback_log:
                callback_log(f"   Resultados FM: {'VÁLIDOS' if datos_fm is not None else 'NULOS/ERROR'}")
                callback_log(f"   Resultados TV: {'VÁLIDOS' if datos_tv is not None else 'NULOS/ERROR'}")
                if datos_fm is not None:
                    callback_log(f"   Filas FM: {len(datos_fm)}")
                if datos_tv is not None:
                    callback_log(f"   Filas TV: {len(datos_tv)}")
            
            # NUEVO: Procesar AM solo para Cuenca
           
            # Permitir continuar si al menos uno de los dos tiene datos
            if datos_fm is None:
                if callback_log:
                    callback_log(f"⚠️  No hay datos FM para {base}, continuando solo con TV")
                # Crear datos FM vacíos pero con estructura correcta
                datos_fm = pd.DataFrame(columns=[
                    "Frecuencia (MHz)", "FECHA DE SUSCRIPCION", "Ocupación (%)",
                    "Level (dBµV/m)", "Bandwidth (Hz)", "Offset (Hz)", "FM (kHz)", "Mes"
                ])

            if datos_tv is None:
                if callback_log:
                    callback_log(f"⚠️  No hay datos TV para {base}, continuando solo con FM")
                # Crear datos TV vacíos pero con estructura correcta
                datos_tv = pd.DataFrame(columns=[
                    "Frecuencia (MHz)", "Banda", "Canal", "Ocupación (%)",
                    "Level (dBµV/m)", "Bandwidth (Hz)", "Offset (Hz)", "AM (%)", "Mes"
                ])
            
            # Verificar que ambos archivos sean del mismo mes
            mes_fm = datos_fm["Mes"].iloc[0] if not datos_fm.empty else None
            mes_tv = datos_tv["Mes"].iloc[0] if not datos_tv.empty else None
            mes_am = datos_am["Mes"].iloc[0] if datos_am is not None and not datos_am.empty else None  # NUEVO

            if mes_fm != mes_tv:
                if callback_log:
                    callback_log(f"⚠️  Los archivos de {base} son de meses diferentes: FM={mes_fm}, TV={mes_tv}")

            # Usar el mes de FM como referencia (o TV si FM no está disponible)
            mes_referencia = mes_fm if mes_fm is not None else mes_tv

            # Crear libro de Excel
            wb = Workbook()

            # Crear hoja para FM
            if "Sheet" in wb.sheetnames:
                ws_fm = wb["Sheet"]
                ws_fm.title = "Datos FM"
            else:
                ws_fm = wb.create_sheet("Datos FM")

            # Variables para frecuencias problemáticas de esta base
            frecuencias_problematicas_fm = []
            frecuencias_problematicas_tv = []
            frecuencias_problematicas_am = []  # NUEVO

            # Formatear hoja FM si hay datos
            if not datos_fm.empty:
                frecuencias_problematicas_fm = formatear_hoja_ocupacion(ws_fm, datos_fm.drop(columns=["Mes"]), "FM", base)
                todas_frecuencias_problematicas["FM"].extend(frecuencias_problematicas_fm)
            else:
                if callback_log:
                    callback_log(f"⚠️  No hay datos FM para {base}")

            # Crear hoja para TV
            ws_tv = wb.create_sheet("Datos TV")
            if not datos_tv.empty:
                frecuencias_problematicas_tv = formatear_hoja_ocupacion(ws_tv, datos_tv.drop(columns=["Mes"]), "TV", base)
                todas_frecuencias_problematicas["TV"].extend(frecuencias_problematicas_tv)
            else:
                if callback_log:
                    callback_log(f"⚠️  No hay datos TV para {base}")

            # NUEVO: Crear hoja para AM (solo para Cuenca)
            if base.lower() == "cuenca" and datos_am is not None and not datos_am.empty:
                ws_am = wb.create_sheet("Datos AM")
                frecuencias_problematicas_am = formatear_hoja_ocupacion_am(ws_am, datos_am.drop(columns=["Mes"]), base)
                todas_frecuencias_problematicas["AM"].extend(frecuencias_problematicas_am)
                if callback_log:
                    callback_log(f"✅ Hoja 'Datos AM' creada para {base}")
            else:
                if callback_log and base.lower() == "cuenca":
                    callback_log(f"⚠️  No hay datos AM para {base}")

            
            
            # Generar nombre de archivo
            codigo_base = obtener_codigo_base(base)
            
            # Normalizar nombre de ciudad
            nombre_ciudad = normalizar_nombre_ciudad(base)
            nombre_mes_completo = obtener_nombre_mes_es(mes_referencia) if mes_referencia else "Desconocido"
            nombre_salida = f"{codigo_base}_Ocupacion{nombre_ciudad}_{nombre_mes_completo}2025.xlsx"


            # Crear hoja "DATOS Manual"
            wb = crear_hoja_datos_manual(wb,nombre_salida)
            
            # Eliminar hoja por defecto si existe
            if "Sheet" in wb.sheetnames and wb.sheetnames[0] == "Sheet":
                del wb["Sheet"]


            ruta_completa = os.path.join(ruta_salida, nombre_salida)
            
            # Guardar archivo
            wb.save(ruta_completa)
            
            # INSERTAR UMBRALES EN EL ARCHIVO EXCEL
            # Determinar qué umbrales usar para esta ciudad
            ciudad_normalizada = normalizar_nombre_ciudad(base)
            if ciudad_normalizada in umbrales:
                umbrales_ciudad = umbrales[ciudad_normalizada]
            else:
                # Usar umbrales globales si no hay específicos para esta ciudad
                umbrales_ciudad = umbrales.get("global", {
                    "FM": 60.0,
                    "AM": 45.0,  # NUEVO
                    "TV": {
                        "tipo": "general",
                        "valor": 45.0,
                        "valores": {
                            "Banda I-III": 47.0,
                            "Banda III": 56.0,
                            "Banda IV-V": 64.0
                        }
                    }
                })
                if callback_log:
                    callback_log(f"⚠️  Usando umbrales globales para {ciudad_normalizada}")
            
            # Insertar umbrales en el archivo Excel
            if insertar_umbrales_excel(ruta_completa, umbrales_ciudad):
                if callback_log:
                    callback_log(f"✅ Umbrales insertados en {nombre_salida}")
            else:
                if callback_log:
                    callback_log(f"⚠️  No se pudieron insertar umbrales en {nombre_salida}")
            
            # Agregar a la lista de archivos generados
            archivos_generados.append(ruta_completa)
            
            # Log de frecuencias problemáticas para esta base
            total_problematicas_base = (len(frecuencias_problematicas_fm) + 
                                      len(frecuencias_problematicas_tv) + 
                                      len(frecuencias_problematicas_am))  # MODIFICADO
            if total_problematicas_base > 0 and callback_log:
                callback_log(f"⚠️  {base}: {total_problematicas_base} frecuencias con ocupación 0%")
            
            if callback_log:
                callback_log(f"✅ Archivo generado: {nombre_salida}")
                
        except Exception as e:
            if callback_log:
                callback_log(f"❌ Error procesando base {base}: {str(e)}")
            import traceback
            traceback.print_exc()
    
    # VERIFICAR SI HAY FRECUENCIAS PROBLEMÁTICAS EN TOTAL
    total_problematicas = (len(todas_frecuencias_problematicas["FM"]) + 
                          len(todas_frecuencias_problematicas["TV"]) + 
                          len(todas_frecuencias_problematicas["AM"]))  # MODIFICADO
    
    if total_problematicas > 0:
        resultado["existen_problemas"] = True
        resultado["datos"] = todas_frecuencias_problematicas
        
        # Guardar advertencia en JSON
        archivo_advertencia = guardar_advertencia_ocupacion_cero(todas_frecuencias_problematicas, ruta_salida)
        if archivo_advertencia:
            if callback_log:
                callback_log(f"📄 Archivo de advertencia generado: {os.path.basename(archivo_advertencia)}")
            # Agregar también el archivo de advertencia a la lista
            archivos_generados.append(archivo_advertencia)
    
    resultado["archivos_generados"] = archivos_generados
    
    if callback_progreso:
        callback_progreso(100)
        
    if callback_log:
        if resultado["existen_problemas"]:
            callback_log(f"✅ Procesamiento completado con {total_problematicas} frecuencias problemáticas")
        else:
            callback_log("✅ Procesamiento completado sin frecuencias problemáticas")
    
    return resultado

# ------------------ EJECUCIÓN DIRECTA (para testing) ------------------

if __name__ == "__main__":
    # Si se ejecuta directamente, usar callbacks simples
    #verificar_encoding_config()
    
    # Luego debug del contenido
    #debug_emisoras_config()


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