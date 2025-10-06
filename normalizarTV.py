import json
import os
from pathlib import Path
from datetime import datetime

def normalizar_frecuencias_tv(archivo_json, callback_log=None):
    """
    Normaliza las frecuencias de TV en un archivo JSON y REEMPLAZA el archivo original:
    - Mantiene frecuencias decimales (que contienen ".")
    - A frecuencias enteras les resta 1.75
    - Para FRECUENCIAS duplicadas, prioriza registros "Activos"
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    # Cargar datos
    with open(archivo_json, 'r', encoding='utf-8') as f:
        datos = json.load(f)
    
    log_message(f"📊 Procesando {len(datos)} registros en {os.path.basename(archivo_json)}")
    
    # Filtrar solo registros de TV
    registros_tv = [registro for registro in datos if registro.get('SERVICIO') == 'TV - Televisión Abierta']
    otros_registros = [registro for registro in datos if registro.get('SERVICIO') != 'TV - Televisión Abierta']
    
    log_message(f"📺 Encontrados {len(registros_tv)} registros de TV para normalizar")
    
    # Procesar cada registro de TV individualmente
    for registro in registros_tv:
        frecuencia = registro.get('FRECUENCIA')
        if frecuencia is not None:
            # Convertir a string para verificar si es decimal
            frecuencia_str = str(frecuencia)
            
            if '.' in frecuencia_str:
                # Mantener frecuencia decimal sin cambios
                log_message(f"🔧 Manteniendo frecuencia decimal: {frecuencia} para {registro.get('RED')}")
            else:
                # Restar 1.75 a frecuencia entera
                try:
                    frecuencia_float = float(frecuencia)
                    nueva_frecuencia = frecuencia_float - 1.75
                    log_message(f"⚙️ Ajustando frecuencia entera: {frecuencia} -> {nueva_frecuencia} para {registro.get('RED')}")
                    registro['FRECUENCIA'] = nueva_frecuencia
                except (ValueError, TypeError):
                    log_message(f"❌ Error procesando frecuencia: {frecuencia} para {registro.get('RED')}")
                    continue
    
    # Ahora eliminar duplicados por FRECUENCIA (después de la normalización)
    registros_tv_sin_duplicados = eliminar_duplicados_por_frecuencia(registros_tv, callback_log)
    
    log_message(f"✅ Registros TV después de eliminar duplicados: {len(registros_tv_sin_duplicados)}")
    
    # Combinar registros procesados de TV con los otros registros
    datos_finales = otros_registros + registros_tv_sin_duplicados
    
    # GUARDAR SOBREESCRIBIENDO EL ARCHIVO ORIGINAL
    with open(archivo_json, 'w', encoding='utf-8') as f:
        json.dump(datos_finales, f, ensure_ascii=False, indent=2)
    
    log_message(f"💾 Procesamiento completado. Archivo ORIGINAL actualizado: {os.path.basename(archivo_json)}")
    
    return datos_finales, archivo_json

def eliminar_duplicados_por_frecuencia(registros_tv, callback_log=None):
    """
    Elimina registros duplicados por FRECUENCIA, priorizando los que están "Activos"
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    registros_por_frecuencia = {}
    
    for registro in registros_tv:
        frecuencia = registro.get('FRECUENCIA')
        
        if frecuencia is None:
            # Si no tiene frecuencia, mantenerlo
            continue
            
        # Crear clave única por FRECUENCIA
        clave = f"{frecuencia}"
        
        if clave not in registros_por_frecuencia:
            registros_por_frecuencia[clave] = []
        registros_por_frecuencia[clave].append(registro)
    
    # Para cada grupo de registros con misma FRECUENCIA, elegir el mejor
    registros_unicos = []
    
    for clave, registros in registros_por_frecuencia.items():
        if len(registros) == 1:
            # Solo un registro, mantenerlo
            registros_unicos.append(registros[0])
        else:
            # Múltiples registros con misma frecuencia, elegir el mejor
            log_message(f"\n🔍 Encontrados {len(registros)} registros para frecuencia: {clave} MHz")
            
            # Mostrar todos los registros encontrados
            for i, r in enumerate(registros, 1):
                estado_est = r.get('ESTADO_ESTACION', 'N/A')
                estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                red = r.get('RED', 'N/A')
                log_message(f"  {i}. {red} - Estados: {estado_est}/{estado_sol}")
            
            # Priorizar registros activos
            registros_activos = [r for r in registros if 
                               r.get('ESTADO_ESTACION') == 'Activo' or 
                               r.get('ESTADO_SOLICITUD') == 'Activo']
            
            if registros_activos:
                # Entre activos, elegir el de mayor vigencia
                registro_elegido = max(registros_activos, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                log_message(f"✅ Elegido registro ACTIVO: {registro_elegido.get('RED')}")
                
                # Mostrar los descartados
                for r in registros:
                    if r != registro_elegido:
                        estado_est = r.get('ESTADO_ESTACION', 'N/A')
                        estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                        log_message(f"  ❌ Descartado: {r.get('RED')} - Estados: {estado_est}/{estado_sol}")
            else:
                # Si no hay activos, elegir el de mayor vigencia entre todos
                registro_elegido = max(registros, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                log_message(f"✅ Elegido registro por VIGENCIA: {registro_elegido.get('RED')}")
    
    return registros_unicos

def convertir_vigencia_a_fecha(vigencia_str):
    """
    Convierte string de vigencia a datetime para comparación
    """
    if not vigencia_str:
        return datetime.min
    
    try:
        # Probar diferentes formatos
        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
            try:
                return datetime.strptime(str(vigencia_str), fmt)
            except ValueError:
                continue
    except:
        pass
    
    return datetime.min

def procesar_todos_los_archivos_en_carpeta(carpeta, callback_log=None, callback_progress=None):
    """
    Procesa todos los archivos JSON en una carpeta y REEMPLAZA los archivos originales
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    def update_progress(value):
        if callback_progress:
            callback_progress(value)
    
    carpeta_path = Path(carpeta)
    archivos_json = list(carpeta_path.glob('*.json'))
    
    log_message(f"📁 Encontrados {len(archivos_json)} archivos para procesar")
    
    resultados = {}
    
    for i, archivo in enumerate(archivos_json):
        log_message(f"\n{'='*50}")
        log_message(f"🔄 PROCESANDO: {archivo.name}")
        log_message(f"{'='*50}")
        
        try:
            datos_normalizados, archivo_salida = normalizar_frecuencias_tv(str(archivo), callback_log)
            resultados[archivo.name] = {
                'archivo_procesado': str(archivo),
                'total_registros': len(datos_normalizados)
            }
            
            # Actualizar progreso
            if callback_progress:
                progreso = 10 + (i / len(archivos_json)) * 80  # 10% a 90%
                update_progress(int(progreso))
                
        except Exception as e:
            error_msg = f"❌ Error procesando {archivo.name}: {str(e)}"
            log_message(error_msg)
            resultados[archivo.name] = {'error': str(e)}
    
    log_message(f"\n{'='*60}")
    log_message(f"✅ PROCESAMIENTO COMPLETADO")
    log_message(f"{'='*60}")
    log_message(f"📊 Total de archivos procesados: {len(resultados)}")
    log_message(f"💾 Todos los archivos originales han sido actualizados")
    
    return resultados

# Función principal para usar desde la GUI
def procesar_normalizacion_tv_desde_gui(callback_log=None, callback_progress=None):
    """
    Función principal para normalizar TV desde la GUI
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    def update_progress(value):
        if callback_progress:
            callback_progress(value)
    
    try:
        log_message("🎯 INICIANDO NORMALIZACIÓN DE FRECUENCIAS TV")
        
        # Configuración
        carpeta_resultados = "SPECTRA_Filtrado"
        
        # Verificar que la carpeta existe
        if not os.path.exists(carpeta_resultados):
            error_msg = f"❌ La carpeta {carpeta_resultados} no existe"
            log_message(error_msg)
            return False
        
        update_progress(5)
        
        # Procesar todos los archivos
        resultados = procesar_todos_los_archivos_en_carpeta(
            carpeta_resultados, 
            callback_log, 
            callback_progress
        )
        
        update_progress(95)
        
        log_message("✅ NORMALIZACIÓN TV COMPLETADA EXITOSAMENTE")
        update_progress(100)
        
        return True
        
    except Exception as e:
        error_msg = f"❌ ERROR en normalización TV: {str(e)}"
        log_message(error_msg)
        return False

# Ejemplo de uso
if __name__ == "__main__":
    # Para uso directo del script
    procesar_normalizacion_tv_desde_gui()