# -*- coding: utf-8 -*-
import os
import json
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ------------------ CONFIGURACIÓN GENERAL ------------------
# Cargar configuración desde archivo
CONFIG_FILE = "config.json"

def cargar_configuracion():
    """Cargar configuración desde archivo JSON"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                return config
        except:
            pass
    
    # Valores por defecto si no hay archivo de configuración
    return {
        "fm_path": "MedicionesFmCSV",
        "tv_path": "MedicionesTvCSV", 
        "ocupacion_output_path": "ReportesOcupacion"
    }

# Cargar configuración al inicio
config = cargar_configuracion()
ruta_fm = config.get("fm_path", "MedicionesFmCSV")
ruta_tv = config.get("tv_path", "MedicionesTvCSV")
ruta_salida = config.get("ocupacion_output_path", "ReportesOcupacion")

# ------------------ FUNCIONES AUXILIARES ------------------

def inicializar_directorios():
    """Crear los directorios necesarios si no existen"""
    os.makedirs(ruta_salida, exist_ok=True)
    os.makedirs(ruta_fm, exist_ok=True)
    os.makedirs(ruta_tv, exist_ok=True)

def obtener_base(nombre_archivo):
    return nombre_archivo.split("_")[0].lower().strip()

def procesar_archivo_fm(ruta_archivo):
    """Procesar archivo FM para extraer los campos requeridos"""
    try:
        # Leer el archivo CSV con low_memory=False para evitar warnings
        df = pd.read_csv(ruta_archivo, encoding="unicode_escape", low_memory=False)
        
        # Filtrar solo las frecuencias hasta 108000000 Hz
        df = df[df["Frecuencia (Hz)"] <= 108000000]
        
        # Seleccionar las columnas requeridas
        columnas_requeridas = ["Frecuencia (Hz)", "FECHA DE SUSCRIPCION", "Level (dBµV/m)", "Bandwidth (Hz)", "Ocupación (%)"]
        
        # Verificar qué columnas existen en el archivo
        columnas_existentes = [col for col in columnas_requeridas if col in df.columns]
        
        # Si no encontramos todas las columnas, intentar con nombres alternativos
        if len(columnas_existentes) < len(columnas_requeridas):
            # Buscar columnas similares
            columnas_disponibles = df.columns.tolist()
            mapeo_columnas = {}
            
            for col_requerida in columnas_requeridas:
                for col_disponible in columnas_disponibles:
                    if col_requerida.lower() in col_disponible.lower():
                        mapeo_columnas[col_requerida] = col_disponible
                        break
            
            # Usar el mapeo para seleccionar las columnas
            if mapeo_columnas:
                df_seleccionado = df[list(mapeo_columnas.values())].copy()
                df_seleccionado.columns = list(mapeo_columnas.keys())
            else:
                # Si no encontramos coincidencias, usar las primeras columnas disponibles
                df_seleccionado = df.iloc[:, :5].copy()
                df_seleccionado.columns = columnas_requeridas[:5]
        else:
            df_seleccionado = df[columnas_existentes].copy()
        
        # Renombrar columnas para consistencia
        df_seleccionado = df_seleccionado.rename(columns={
            "Frecuencia (Hz)": "Frecuencia (Hz)",
            "FECHA DE SUSCRIPCION": "FECHA DE SUSCRIPCION",
            "Level (dBµV/m)": "Level Media (dBµV/m)",
            "Bandwidth (Hz)": "Bandwidth Media (Hz)",
            "Ocupación (%)": "Ocupación (%)"
        })
        
        return df_seleccionado
        
    except Exception as e:
        print(f"Error al procesar archivo FM: {str(e)}")
        return None

def procesar_archivo_tv(ruta_archivo):
    """Procesar archivo TV para extraer los campos requeridos"""
    try:
        # Leer el archivo CSV con low_memory=False para evitar warnings
        df = pd.read_csv(ruta_archivo, encoding="unicode_escape", low_memory=False)
        
        # Filtrar solo las frecuencias hasta 108000000 Hz
        df = df[df["Frecuencia (Hz)"] <= 108000000]
        
        # Para TV, necesitamos determinar qué campos extraer
        columnas_disponibles = df.columns.tolist()
        
        # Buscar columnas relevantes
        columnas_seleccionadas = []
        nombres_finales = []
        
        # Frecuencia
        for col in columnas_disponibles:
            if "frecuencia" in col.lower():
                columnas_seleccionadas.append(col)
                nombres_finales.append("Frecuencia (Hz)")
                break
        
        # Fecha de suscripción
        for col in columnas_disponibles:
            if "suscripcion" in col.lower() or "fecha" in col.lower():
                columnas_seleccionadas.append(col)
                nombres_finales.append("FECHA DE SUSCRIPCION")
                break
        
        # Level
        for col in columnas_disponibles:
            if "level" in col.lower():
                columnas_seleccionadas.append(col)
                nombres_finales.append("Level Media (dBµV/m)")
                break
        
        # Bandwidth
        for col in columnas_disponibles:
            if "bandwidth" in col.lower():
                columnas_seleccionadas.append(col)
                nombres_finales.append("Bandwidth Media (Hz)")
                break
        
        # Ocupación
        for col in columnas_disponibles:
            if "ocupación" in col.lower() or "ocupacion" in col.lower():
                columnas_seleccionadas.append(col)
                nombres_finales.append("Ocupación (%)")
                break
        
        # Si no encontramos todas las columnas, usar las primeras disponibles
        if len(columnas_seleccionadas) < 5:
            columnas_seleccionadas = columnas_disponibles[:5]
            nombres_finales = ["Frecuencia (Hz)", "FECHA DE SUSCRIPCION", "Level Media (dBµV/m)", 
                              "Bandwidth Media (Hz)", "Ocupación (%)"][:len(columnas_seleccionadas)]
        
        df_seleccionado = df[columnas_seleccionadas].copy()
        df_seleccionado.columns = nombres_finales
        
        return df_seleccionado
        
    except Exception as e:
        print(f"Error al procesar archivo TV: {str(e)}")
        return None

def crear_archivo_excel(datos_fm, datos_tv, nombre_archivo):
    """Crear archivo Excel con los datos procesados"""
    try:
        print(f"Intentando crear archivo: {nombre_archivo}")
        
        # VERIFICAR TAMAÑO DE DATOS
        if datos_fm is not None:
            print(f"Datos FM: {len(datos_fm)} filas")
            if len(datos_fm) > 1000000:
                print("⚠️ ADVERTENCIA: Datos FM exceden el límite de Excel")
                # Limitar a las primeras 1,000,000 filas
                datos_fm = datos_fm.head(1000000)
        
        if datos_tv is not None:
            print(f"Datos TV: {len(datos_tv)} filas")
            if len(datos_tv) > 1000000:
                print("⚠️ ADVERTENCIA: Datos TV exceden el límite de Excel")
                # Limitar a las primeras 1,000,000 filas
                datos_tv = datos_tv.head(1000000)
        
        wb = Workbook()
        
        # Hoja de procesamiento FM
        if datos_fm is not None and not datos_fm.empty:
            ws_fm = wb.active
            ws_fm.title = "Procesamiento FM"
            
            # Agregar datos
            for i, row in enumerate(dataframe_to_rows(datos_fm, index=False, header=True), 1):
                for j, value in enumerate(row, 1):
                    ws_fm.cell(row=i, column=j, value=value)
            
            # Ajustar anchos de columnas
            for column in ws_fm.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws_fm.column_dimensions[column_letter].width = adjusted_width
            
            # Formato de encabezados
            for cell in ws_fm[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
        else:
            # Crear hoja vacía con mensaje
            ws_fm = wb.active
            ws_fm.title = "Procesamiento FM"
            ws_fm.cell(row=1, column=1, value="No hay datos de FM disponibles")
        
        # Hoja de procesamiento TV
        if datos_tv is not None and not datos_tv.empty:
            if "Procesamiento TV" not in [sheet.title for sheet in wb.worksheets]:
                ws_tv = wb.create_sheet("Procesamiento TV")
            
            # Agregar datos
            for i, row in enumerate(dataframe_to_rows(datos_tv, index=False, header=True), 1):
                for j, value in enumerate(row, 1):
                    ws_tv.cell(row=i, column=j, value=value)
            
            # Ajustar anchos de columnas
            for column in ws_tv.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws_tv.column_dimensions[column_letter].width = adjusted_width
            
            # Formato de encabezados
            for cell in ws_tv[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
        else:
            if "Procesamiento TV" not in [sheet.title for sheet in wb.worksheets]:
                ws_tv = wb.create_sheet("Procesamiento TV")
                ws_tv.cell(row=1, column=1, value="No hay datos de TV disponibles")
        
        # Guardar archivo
        print("Guardando archivo Excel...")
        wb.save(nombre_archivo)
        print("Archivo Excel guardado exitosamente")
        return True
        
    except Exception as e:
        print(f"Error al crear archivo Excel: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return False

def crear_archivos_csv(datos_fm, datos_tv, ruta_base):
    """Crear archivos CSV en lugar de Excel para datos muy grandes"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if datos_fm is not None and not datos_fm.empty:
            nombre_archivo = os.path.join(ruta_base, f"Reporte_Ocupacion_FM_{timestamp}.csv")
            datos_fm.to_csv(nombre_archivo, index=False, encoding='utf-8')
            print(f"Archivo CSV FM creado: {nombre_archivo}")
        
        if datos_tv is not None and not datos_tv.empty:
            nombre_archivo = os.path.join(ruta_base, f"Reporte_Ocupacion_TV_{timestamp}.csv")
            datos_tv.to_csv(nombre_archivo, index=False, encoding='utf-8')
            print(f"Archivo CSV TV creado: {nombre_archivo}")
        
        return True
        
    except Exception as e:
        print(f"Error al crear archivos CSV: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return False

# ------------------ FUNCIÓN PRINCIPAL DE OCUPACIÓN ------------------

def procesar_ocupacion(callback_progreso=None, callback_log=None):
    """
    Función principal para procesar datos de ocupación
    """
    # Inicializar directorios
    inicializar_directorios()
    
    if callback_log:
        callback_log("Iniciando procesamiento de ocupación...")
        callback_log(f"Ruta FM: {ruta_fm}")
        callback_log(f"Ruta TV: {ruta_tv}")
        callback_log(f"Ruta salida: {ruta_salida}")
    
    # Obtener listas de archivos
    try:
        archivos_fm = [f for f in os.listdir(ruta_fm) if f.endswith(".csv")]
        archivos_tv = [f for f in os.listdir(ruta_tv) if f.endswith(".csv")]
        
        if callback_log:
            callback_log(f"Encontrados {len(archivos_fm)} archivos FM y {len(archivos_tv)} archivos TV")
    except Exception as e:
        if callback_log:
            callback_log(f"Error al leer archivos: {str(e)}")
        return False

    # Procesar cada archivo FM
    datos_fm_combinados = pd.DataFrame()
    for i, archivo in enumerate(archivos_fm):
        if callback_progreso:
            progreso = int((i / len(archivos_fm)) * 50)  # Primera mitad para FM
            callback_progreso(progreso)
            
        if callback_log:
            callback_log(f"Procesando archivo FM: {archivo}")
        
        try:
            ruta_completa = os.path.join(ruta_fm, archivo)
            datos_fm = procesar_archivo_fm(ruta_completa)
            
            if datos_fm is not None and not datos_fm.empty:
                # Agregar columna con el nombre del archivo
                datos_fm["Archivo"] = archivo
                
                # Combinar con datos anteriores
                datos_fm_combinados = pd.concat([datos_fm_combinados, datos_fm], ignore_index=True)
                
                if callback_log:
                    callback_log(f"✅ Archivo FM procesado: {archivo}")
            else:
                if callback_log:
                    callback_log(f"⚠️  Archivo FM sin datos: {archivo}")
                    
        except Exception as e:
            if callback_log:
                callback_log(f"❌ Error procesando archivo FM {archivo}: {str(e)}")
    
    # Procesar cada archivo TV
    datos_tv_combinados = pd.DataFrame()
    for i, archivo in enumerate(archivos_tv):
        if callback_progreso:
            progreso = 50 + int((i / len(archivos_tv)) * 50)  # Segunda mitad para TV
            callback_progreso(progreso)
            
        if callback_log:
            callback_log(f"Procesando archivo TV: {archivo}")
        
        try:
            ruta_completa = os.path.join(ruta_tv, archivo)
            datos_tv = procesar_archivo_tv(ruta_completa)
            
            if datos_tv is not None and not datos_tv.empty:
                # Agregar columna con el nombre del archivo
                datos_tv["Archivo"] = archivo
                
                # Combinar con datos anteriores
                datos_tv_combinados = pd.concat([datos_tv_combinados, datos_tv], ignore_index=True)
                
                if callback_log:
                    callback_log(f"✅ Archivo TV procesado: {archivo}")
            else:
                if callback_log:
                    callback_log(f"⚠️  Archivo TV sin datos: {archivo}")
                    
        except Exception as e:
            if callback_log:
                callback_log(f"❌ Error procesando archivo TV {archivo}: {str(e)}")
    
    # Verificar tamaño total de datos
    total_filas = 0
    if datos_fm_combinados is not None:
        total_filas += len(datos_fm_combinados)
    if datos_tv_combinados is not None:
        total_filas += len(datos_tv_combinados)
    
    if callback_log:
        callback_log(f"Total de filas a procesar: {total_filas}")
    
    # Crear archivo de salida
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_base = ruta_salida
        
        # Decidir qué formato usar basado en el tamaño
        if total_filas > 2000000:  # Si más de 2 millones de filas, usar CSV
            if callback_log:
                callback_log("Datos muy grandes, usando formato CSV...")
            exito = crear_archivos_csv(datos_fm_combinados, datos_tv_combinados, ruta_base)
        else:  # Usar Excel normal
            if callback_log:
                callback_log("Creando archivo Excel con los datos procesados...")
            nombre_archivo = os.path.join(ruta_base, f"Reporte_Ocupacion_{timestamp}.xlsx")
            exito = crear_archivo_excel(datos_fm_combinados, datos_tv_combinados, nombre_archivo)
        
        if exito:
            if callback_log:
                callback_log(f"✅ Archivo creado exitosamente")
        else:
            if callback_log:
                callback_log("❌ Error al crear archivo")
            return False
            
    except Exception as e:
        if callback_log:
            callback_log(f"❌ Error al crear archivo: {str(e)}")
        return False
    
    if callback_progreso:
        callback_progreso(100)
        
    if callback_log:
        callback_log("Procesamiento de ocupación completado")
    
    return True

# ------------------ EJECUCIÓN DIRECTA (para testing) ------------------

if __name__ == "__main__":
    # Si se ejecuta directamente, usar callbacks simples
    def mostrar_progreso(progreso):
        print(f"Progreso: {progreso}%")
    
    def mostrar_log(mensaje):
        print(mensaje)
    
    # Procesar datos de ocupación
    resultado = procesar_ocupacion(mostrar_progreso, mostrar_log)
    
    if resultado:
        print("Procesamiento de ocupación completado con éxito")
    else:
        print("Ocurrieron errores durante el procesamiento de ocupación")