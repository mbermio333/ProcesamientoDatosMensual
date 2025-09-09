# -*- coding: utf-8 -*-
import os
import json
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# ------------------ CONFIGURACIÓN GENERAL ------------------
# Cargar configuración desde archivo
CONFIG_FILE = "config.json"

def cargar_configuracion():
    """Cargar configuración desde archivo JSON"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
        except:
            pass
    
    # Valores por defecto si no hay archivo de configuración
    return {
        "fm_path": "MedicionesFmCSV",
        "tv_path": "MedicionesTvCSV", 
        "output_path": "ReportesOcupacion"
    }

# Cargar configuración al inicio
config = cargar_configuracion()
ruta_fm = config.get("fm_path", "MedicionesFmCSV")
ruta_tv = config.get("tv_path", "MedicionesTvCSV")
ruta_salida = config.get("output_path", "ReportesOcupacion")
fecha_actual = datetime.now().strftime("%d/%m/%Y")

# ------------------ FUNCIONES AUXILIARES ------------------

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
        "canar": "SCS-L03",
        "macas": "SCS-L04",
        "machala": "SCC-L04",
        "cuenca": "SCS-L05"
    }
    return correspondencia.get(base.lower(), f"SCS-{base.upper()}")

def obtener_nombre_mes_es(numero_mes):
    """Convierte el número de mes a nombre en español"""
    meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    return meses[numero_mes - 1] if 1 <= numero_mes <= 12 else ""

def obtener_base(nombre_archivo):
    """Obtiene el nombre base del archivo"""
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

def formatear_hoja_ocupacion(ws, datos, tipo):
    """Formatea una hoja de ocupación con bordes y estilos"""
    # Limpiar hoja existente
    ws.delete_rows(1, ws.max_row)
    
    # Agregar encabezado
    if tipo == "FM":
        encabezados = ["Frecuencia (MHz)", "FECHA DE SUSCRIPCION", "Ocupación (%)"]
    else:  # TV
        encabezados = ["Frecuencia (MHz)", "Canal", "Ocupación (%)"]
    
    # Escribir encabezados
    for col, encabezado in enumerate(encabezados, 1):
        ws.cell(row=1, column=col, value=encabezado)
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")
    
    # Escribir datos
    for fila_idx, (_, fila) in enumerate(datos.iterrows(), 2):
        if tipo == "FM":
            ws.cell(row=fila_idx, column=1, value=fila["Frecuencia (MHz)"])
            ws.cell(row=fila_idx, column=2, value=fila["FECHA DE SUSCRIPCION"])
            # Conservar los decimales originales de ocupación
            ws.cell(row=fila_idx, column=3, value=fila["Ocupación (%)"])
        else:  # TV
            ws.cell(row=fila_idx, column=1, value=fila["Frecuencia (MHz)"])
            ws.cell(row=fila_idx, column=2, value=fila["Canal"])
            # Conservar los decimales originales de ocupación
            ws.cell(row=fila_idx, column=3, value=fila["Ocupación (%)"])
    
    # Aplicar bordes y formato
    thin = Side(border_style="thin")
    borde_grueso = Side(border_style="medium")
    
    for row in range(1, len(datos) + 2):
        for col in range(1, len(encabezados) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            if row == 1:  # Encabezados
                cell.border = Border(top=borde_grueso, bottom=thin,
                                   left=borde_grueso if col == 1 else thin,
                                   right=borde_grueso if col == len(encabezados) else thin)
            elif row == len(datos) + 1:  # Última fila
                cell.border = Border(top=thin, bottom=borde_grueso,
                                   left=borde_grueso if col == 1 else thin,
                                   right=borde_grueso if col == len(encabezados) else thin)
            elif col == 1:  # Primera columna
                cell.border = Border(left=borde_grueso, top=thin, bottom=thin, right=thin)
            elif col == len(encabezados):  # Última columna
                cell.border = Border(right=borde_grueso, top=thin, bottom=thin, left=thin)
    
    # Ajustar anchos de columnas
    for col in range(1, len(encabezados) + 1):
        col_letter = get_column_letter(col)
        max_length = 0
        for cell in ws[col_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

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

def procesar_archivo_fm(ruta_archivo):
    """Procesa archivo FM y extrae datos de ocupación en el rango desde 88.1 MHz"""
    try:
        # Primero reducir el archivo
        if not reducir_archivo_csv(ruta_archivo, "FM"):
            return None
            
        # Leer el archivo reducido
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Filtrar por rango de frecuencia (desde 88.1 MHz)
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        df_filtrado = df[(df["Frecuencia (MHz)"] >= 88.1)]
        
        if df_filtrado.empty:
            return None
        
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
        
        # Buscar la columna de ocupación (puede tener diferentes nombres por encoding)
        columna_ocupacion = None
        for col in df_filtrado.columns:
            if 'ocupaci' in col.lower():
                columna_ocupacion = col
                break
        
        # Calcular ocupación usando los valores reales del archivo
        ocupacion_data = []
        
        if columna_ocupacion:
            # Usar los valores reales de ocupación del archivo
            for _, row in df_filtrado.iterrows():
                try:
                    # Limpiar y convertir el valor de ocupación
                    ocupacion_val = limpiar_valor_numerico(row[columna_ocupacion])
                    if not np.isnan(ocupacion_val):
                        # Conservar el valor original con sus decimales
                        ocupacion_data.append({
                            "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                            "FECHA DE SUSCRIPCION": datetime.now().strftime("%Y-%m-%d"),
                            "Ocupación (%)": ocupacion_val  # Valor original con decimales
                        })
                except (ValueError, TypeError):
                    continue
        else:
            # Si no hay columna de ocupación, usar 100% para frecuencias con señal
            frecuencias_unicas = df_filtrado["Frecuencia (MHz)"].unique()
            
            for freq in frecuencias_unicas:
                # Calcular porcentaje de tiempo con señal en esta frecuencia
                mediciones_freq = df_filtrado[df_filtrado["Frecuencia (MHz)"] == freq]
                if not mediciones_freq.empty:
                    # Asumir que si hay medición, hay 100% de ocupación
                    ocupacion_data.append({
                        "Frecuencia (MHz)": freq,
                        "FECHA DE SUSCRIPCION": datetime.now().strftime("%Y-%m-%d"),
                        "Ocupación (%)": 100.0
                    })
        
        resultado = pd.DataFrame(ocupacion_data)
        resultado["Mes"] = mes_objetivo
        return resultado
        
    except Exception as e:
        print(f"Error procesando archivo FM {ruta_archivo}: {e}")
        import traceback
        traceback.print_exc()
        return None

def procesar_archivo_tv(ruta_archivo):
    """Procesa archivo TV y extrae datos de ocupación en el rango 55.25-693.25 MHz"""
    try:
        # Primero reducir el archivo
        if not reducir_archivo_csv(ruta_archivo, "TV"):
            return None
            
        # Leer el archivo reducido
        df = pd.read_csv(ruta_archivo, encoding='latin-1', low_memory=False)
        
        # Limpiar nombres de columnas
        df.columns = [col.strip().replace('Ą', 'u').replace('¾', 'o') for col in df.columns]
        
        # Filtrar por rango de frecuencia correcto (55.25-693.25 MHz)
        df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"].apply(limpiar_valor_numerico) / 1_000_000
        df_filtrado = df[(df["Frecuencia (MHz)"] >= 55.25) & (df["Frecuencia (MHz)"] <= 693.25)]
        
        if df_filtrado.empty:
            return None
        
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
        
        # Mapeo de frecuencia a canal (para TV)
        def frecuencia_a_canal(freq):
            # Bandas de TV según estándares internacionales
            if 55.25 <= freq <= 88: return "Bandas I-III (VHF)"
            elif 174 <= freq <= 216: return "Banda III (VHF)"
            elif 470 <= freq <= 608: return "Bandas IV-V (UHF)"
            elif 614 <= freq <= 698: return "Bandas IV-V (UHF)"
            else: return "Otra banda"
        
        # Buscar la columna de ocupación (puede tener diferentes nombres por encoding)
        columna_ocupacion = None
        for col in df_filtrado.columns:
            if 'ocupaci' in col.lower():
                columna_ocupacion = col
                break
        
        # Calcular ocupación usando los valores reales del archivo
        ocupacion_data = []
        
        if columna_ocupacion:
            # Usar los valores reales de ocupación del archivo
            for _, row in df_filtrado.iterrows():
                try:
                    # Limpiar y convertir el valor de ocupación
                    ocupacion_val = limpiar_valor_numerico(row[columna_ocupacion])
                    if not np.isnan(ocupacion_val):
                        # Conservar el valor original con sus decimales
                        ocupacion_data.append({
                            "Frecuencia (MHz)": row["Frecuencia (MHz)"],
                            "Canal": frecuencia_a_canal(row["Frecuencia (MHz)"]),
                            "Ocupación (%)": ocupacion_val  # Valor original con decimales
                        })
                except (ValueError, TypeError):
                    continue
        else:
            # Si no hay columna de ocupación, usar 100% para frecuencias con señal
            frecuencias_unicas = df_filtrado["Frecuencia (MHz)"].unique()
            
            for freq in frecuencias_unicas:
                # Calcular porcentaje de tiempo con señal en esta frecuencia
                mediciones_freq = df_filtrado[df_filtrado["Frecuencia (MHz)"] == freq]
                if not mediciones_freq.empty:
                    # Asumir que si hay medición, hay 100% de ocupación
                    ocupacion_data.append({
                        "Frecuencia (MHz)": freq,
                        "Canal": frecuencia_a_canal(freq),
                        "Ocupación (%)": 100.0
                    })
        
        resultado = pd.DataFrame(ocupacion_data)
        resultado["Mes"] = mes_objetivo
        return resultado
        
    except Exception as e:
        print(f"Error procesando archivo TV {ruta_archivo}: {e}")
        import traceback
        traceback.print_exc()
        return None

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
            # Procesar archivos FM y TV
            datos_fm = procesar_archivo_fm(archivos_fm[base])
            datos_tv = procesar_archivo_tv(archivos_tv[base])
            
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
            
            if not datos_fm.empty:
                formatear_hoja_ocupacion(ws_fm, datos_fm.drop(columns=["Mes"]), "FM")
            
            # Crear hoja para TV
            ws_tv = wb.create_sheet("Datos TV")
            if not datos_tv.empty:
                formatear_hoja_ocupacion(ws_tv, datos_tv.drop(columns=["Mes"]), "TV")
            
            # Eliminar hoja por defecto si existe
            if "Sheet" in wb.sheetnames and wb.sheetnames[0] == "Sheet":
                del wb["Sheet"]
            
            # Generar nombre de archivo
            codigo_base = obtener_codigo_base(base)
            nombre_ciudad = "TAMBO" if base.lower() == "canar" else base.upper()
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