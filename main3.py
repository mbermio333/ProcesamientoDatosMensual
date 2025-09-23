# -*- coding: utf-8 -*-
import os
import pandas as pd
import numpy as np
from datetime import datetime
import json

def obtener_frecuencias_observacion(ruta_archivo_excel, ciudad):
    """
    Obtiene las frecuencias en observación (con "No identificada" en la columna Estación)
    tanto para FM como para TV de un archivo Excel generado en la pestaña de ocupación
    """
    try:
        # Verificar si el archivo existe
        if not os.path.exists(ruta_archivo_excel):
            return {"FM": [], "TV": [], "error": f"Archivo no encontrado: {ruta_archivo_excel}"}
        
        # Leer el archivo Excel
        datos_fm = []
        datos_tv = []
        
        # Procesar hoja FM
        try:
            df_fm = pd.read_excel(ruta_archivo_excel, sheet_name='Datos FM')
            
            # Buscar filas con "No identificada" en la columna Estación
            if 'Estación' in df_fm.columns:
                filas_no_identificadas = df_fm[df_fm['Estación'].str.contains('No identificada', na=False, case=False)]
                
                for _, fila in filas_no_identificadas.iterrows():
                    datos_fm.append({
                        'Frecuencia (MHz)': fila.get('Frecuencia (MHz)', ''),
                        'Estación': fila.get('Estación', ''),
                        'Ocupación (%)': fila.get('Ocupación (%)', ''),
                        'Level (dBµV/m)': fila.get('Level (dBµV/m)', ''),
                        'Bandwidth (Hz)': fila.get('Bandwidth (Hz)', ''),
                        'Offset (Hz)': fila.get('Offset (Hz)', ''),
                        'FM (kHz)': fila.get('FM (kHz)', ''),
                        'Tipo': 'FM'
                    })
        except Exception as e:
            print(f"Error procesando hoja FM: {e}")
        
        # Procesar hoja TV
        try:
            df_tv = pd.read_excel(ruta_archivo_excel, sheet_name='Datos TV')
            
            # Buscar filas con "No identificada" en la columna Estación
            if 'Estación' in df_tv.columns:
                filas_no_identificadas = df_tv[df_tv['Estación'].str.contains('No identificada', na=False, case=False)]
                
                for _, fila in filas_no_identificadas.iterrows():
                    datos_tv.append({
                        'Frecuencia (MHz)': fila.get('Frecuencia (MHz)', ''),
                        'Estación': fila.get('Estación', ''),
                        'Banda': fila.get('Banda', ''),
                        'Canal': fila.get('Canal', ''),
                        'Ocupación (%)': fila.get('Ocupación (%)', ''),
                        'Level (dBµV/m)': fila.get('Level (dBµV/m)', ''),
                        'Bandwidth (Hz)': fila.get('Bandwidth (Hz)', ''),
                        'Offset (Hz)': fila.get('Offset (Hz)', ''),
                        'AM (%)': fila.get('AM (%)', ''),
                        'Tipo': 'TV'
                    })
        except Exception as e:
            print(f"Error procesando hoja TV: {e}")
        
        return {
            "FM": datos_fm,
            "TV": datos_tv,
            "total_fm": len(datos_fm),
            "total_tv": len(datos_tv),
            "archivo": os.path.basename(ruta_archivo_excel)
        }
        
    except Exception as e:
        return {"FM": [], "TV": [], "error": f"Error general: {str(e)}"}

def buscar_archivos_ocupacion_ciudad(ruta_salida_ocupacion, ciudad):
    """
    Busca archivos Excel de ocupación para una ciudad específica
    """
    try:
        if not os.path.exists(ruta_salida_ocupacion):
            return []
        
        archivos_encontrados = []
        
        # Normalizar nombre de ciudad para búsqueda
        ciudad_normalizada = ciudad.upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        
        for archivo in os.listdir(ruta_salida_ocupacion):
            if archivo.endswith('.xlsx') and ciudad_normalizada in archivo.upper():
                archivos_encontrados.append({
                    'nombre': archivo,
                    'ruta_completa': os.path.join(ruta_salida_ocupacion, archivo),
                    'fecha_modificacion': datetime.fromtimestamp(os.path.getmtime(os.path.join(ruta_salida_ocupacion, archivo)))
                })
        
        # Ordenar por fecha de modificación (más reciente primero)
        archivos_encontrados.sort(key=lambda x: x['fecha_modificacion'], reverse=True)
        
        return archivos_encontrados
        
    except Exception as e:
        print(f"Error buscando archivos: {e}")
        return []

def obtener_todas_frecuencias_observacion(ruta_salida_ocupacion, ciudad):
    """
    Obtiene todas las frecuencias en observación para una ciudad específica
    """
    try:
        archivos = buscar_archivos_ocupacion_ciudad(ruta_salida_ocupacion, ciudad)
        
        if not archivos:
            return {"FM": [], "TV": [], "error": f"No se encontraron archivos para la ciudad: {ciudad}"}
        
        # Usar el archivo más reciente
        archivo_mas_reciente = archivos[0]
        
        resultado = obtener_frecuencias_observacion(archivo_mas_reciente['ruta_completa'], ciudad)
        resultado['archivo_utilizado'] = archivo_mas_reciente['nombre']
        resultado['fecha_archivo'] = archivo_mas_reciente['fecha_modificacion'].strftime("%Y-%m-%d %H:%M:%S")
        
        return resultado
        
    except Exception as e:
        return {"FM": [], "TV": [], "error": f"Error obteniendo frecuencias: {str(e)}"}

# Función principal para ser llamada desde la GUI
def procesar_frecuencias_observacion(ruta_salida_ocupacion, ciudad, callback_log=None):
    """
    Función principal para procesar frecuencias en observación
    """
    if callback_log:
        callback_log(f"Buscando frecuencias en observación para: {ciudad}")
        callback_log(f"Ruta de búsqueda: {ruta_salida_ocupacion}")
    
    resultado = obtener_todas_frecuencias_observacion(ruta_salida_ocupacion, ciudad)
    
    if callback_log:
        if resultado.get('error'):
            callback_log(f"❌ Error: {resultado['error']}")
        else:
            callback_log(f"✅ Encontradas {resultado['total_fm']} frecuencias FM y {resultado['total_tv']} frecuencias TV en observación")
            callback_log(f"📊 Archivo utilizado: {resultado.get('archivo_utilizado', 'N/A')}")
    
    return resultado

if __name__ == "__main__":
    # Ejemplo de uso
    def ejemplo_callback(mensaje):
        print(mensaje)
    
    # Ejemplo de prueba
    resultado = procesar_frecuencias_observacion(
        ruta_salida_ocupacion="ReportesOcupacion",
        ciudad="ZAMORA",
        callback_log=ejemplo_callback
    )
    
    print("\nResultado completo:")
    print(json.dumps(resultado, indent=2, ensure_ascii=False, default=str))