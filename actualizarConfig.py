import json
import os
from pathlib import Path
import shutil

def actualizar_config_desde_spectra(config_path, spectra_folder, callback_log=None, callback_progress=None):
    """
    Actualiza el archivo config.json con los datos de los archivos Spectra
    PRESERVANDO las emisoras manuales "No identificada_OBSERVACION"
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
        # Cargar el archivo config.json existente
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        log_message(f"📂 Cargado config.json con {len(config['emisoras_por_ciudad'])} ciudades")
        
        # Mapeo de nombres de archivos spectra a ciudades en config
        mapeo_ciudades = {
            'cuenca_spectra.json': 'cuenca',
            'loja_spectra.json': 'loja',
            'zamora_spectra.json': 'zamora', 
            'macas_spectra.json': 'macas',
            'machala_spectra.json': 'machala',
            'tambo_spectra.json': 'tambo'
        }
        
        total_ciudades = len(mapeo_ciudades)
        ciudades_procesadas = 0
        
        # Procesar cada archivo Spectra
        for archivo_spectra, ciudad in mapeo_ciudades.items():
            ruta_spectra = os.path.join(spectra_folder, archivo_spectra)
            
            if os.path.exists(ruta_spectra):
                log_message(f"\n{'='*50}")
                log_message(f"🔄 Procesando: {archivo_spectra} -> Ciudad: {ciudad}")
                log_message(f"{'='*50}")
                
                try:
                    # Cargar archivo Spectra
                    with open(ruta_spectra, 'r', encoding='utf-8') as f:
                        datos_spectra = json.load(f)
                    
                    # Actualizar config para esta ciudad (FUSIONANDO, no reemplazando)
                    actualizar_ciudad_config_fusion(config, ciudad, datos_spectra, callback_log)
                    
                    ciudades_procesadas += 1
                    progreso = 10 + (ciudades_procesadas / total_ciudades) * 80
                    update_progress(int(progreso))
                    
                except Exception as e:
                    error_msg = f"❌ Error procesando {archivo_spectra}: {str(e)}"
                    log_message(error_msg)
            else:
                log_message(f"⚠️ Archivo no encontrado: {ruta_spectra}")
        
        # Guardar el config actualizado
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        
        log_message(f"\n{'='*60}")
        log_message(f"✅ ACTUALIZACIÓN COMPLETADA")
        log_message(f"{'='*60}")
        log_message(f"💾 Archivo config.json actualizado: {config_path}")
        
        update_progress(100)
        
        return config
        
    except Exception as e:
        error_msg = f"❌ ERROR en actualización de configuración: {str(e)}"
        log_message(error_msg)
        return None

def actualizar_ciudad_config_fusion(config, ciudad, datos_spectra, callback_log=None):
    """
    Actualiza los datos de una ciudad específica FUSIONANDO con las existentes
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    if ciudad not in config['emisoras_por_ciudad']:
        log_message(f"⚠️ Ciudad '{ciudad}' no encontrada en config.json, creando...")
        config['emisoras_por_ciudad'][ciudad] = {'FM': [], 'TV': [], 'AM': []}
    
    # Obtener las emisoras actuales (para preservar las manuales)
    emisoras_actuales = config['emisoras_por_ciudad'][ciudad]
    emisoras_fm_actuales = emisoras_actuales.get('FM', [])
    emisoras_tv_actuales = emisoras_actuales.get('TV', [])
    emisoras_am_actuales = emisoras_actuales.get('AM', [])
    
    # Separar emisoras manuales de las automáticas
    emisoras_manuales_fm = [e for e in emisoras_fm_actuales if 'No identificada_OBSERVACION' in e.get('nombre', '')]
    emisoras_manuales_tv = [e for e in emisoras_tv_actuales if 'No identificada_OBSERVACION' in e.get('nombre', '')]
    emisoras_manuales_am = [e for e in emisoras_am_actuales if 'No identificada_OBSERVACION' in e.get('nombre', '')]
    
    log_message(f"📝 Emisoras manuales encontradas: FM({len(emisoras_manuales_fm)}), TV({len(emisoras_manuales_tv)}), AM({len(emisoras_manuales_am)})")
    
    # Procesar datos de SPECTRA
    emisoras_spectra_fm = []
    emisoras_spectra_tv = []
    emisoras_spectra_am = []
    
    # Contadores para estadísticas
    total_procesadas = 0
    total_filtradas_fm = 0
    
    # Procesar cada registro del Spectra
    for registro in datos_spectra:
        servicio = registro.get('SERVICIO', '')
        red = registro.get('RED', '')
        frecuencia = registro.get('FRECUENCIA')
        estado_estacion = registro.get('ESTADO_ESTACION', '')
        estado_solicitud = registro.get('ESTADO_SOLICITUD', '')
        
        # Validar datos requeridos
        if not servicio or not red or frecuencia is None:
            continue
        
        # Mapear servicio a tipo
        tipo = mapear_servicio_a_tipo(servicio)
        if not tipo:
            continue
        
        # FILTRO: Excluir FM con frecuencia > 108 MHz
        if tipo == 'FM' and frecuencia > 108:
            log_message(f"  🗑️ Filtrada FM > 108MHz: {red} - {frecuencia} MHz")
            total_filtradas_fm += 1
            continue
        
        # Formatear nombre según estado
        nombre_formateado = formatear_nombre_emisora(red, estado_estacion, estado_solicitud)
        
        # Crear objeto emisora
        emisora = {
            "nombre": nombre_formateado,
            "frecuencia": frecuencia,
            "tipo": tipo
        }
        
        # Agregar a la lista correspondiente
        if tipo == 'FM':
            emisoras_spectra_fm.append(emisora)
        elif tipo == 'TV':
            emisoras_spectra_tv.append(emisora)
        elif tipo == 'AM':
            emisoras_spectra_am.append(emisora)
        
        total_procesadas += 1
    
    # Eliminar duplicados por frecuencia (mantener el último encontrado)
    emisoras_spectra_fm = eliminar_duplicados_por_frecuencia(emisoras_spectra_fm)
    emisoras_spectra_tv = eliminar_duplicados_por_frecuencia(emisoras_spectra_tv)
    emisoras_spectra_am = eliminar_duplicados_por_frecuencia(emisoras_spectra_am)
    
    # FUSIONAR: Combinar emisoras de SPECTRA con emisoras manuales
    # Para FM y TV: priorizar SPECTRA, pero mantener manuales que no entren en conflicto
    emisoras_fm_final = fusionar_emisoras(emisoras_spectra_fm, emisoras_manuales_fm, callback_log)
    emisoras_tv_final = fusionar_emisoras(emisoras_spectra_tv, emisoras_manuales_tv, callback_log)
    emisoras_am_final = fusionar_emisoras(emisoras_spectra_am, emisoras_manuales_am, callback_log)
    
    # Actualizar config
    config['emisoras_por_ciudad'][ciudad]['FM'] = emisoras_fm_final
    config['emisoras_por_ciudad'][ciudad]['TV'] = emisoras_tv_final
    config['emisoras_por_ciudad'][ciudad]['AM'] = emisoras_am_final
    
    # Estadísticas
    log_message(f"📊 Procesadas: {total_procesadas} emisoras de SPECTRA")
    log_message(f"📻 FM: {len(emisoras_spectra_fm)} de SPECTRA + {len(emisoras_manuales_fm)} manuales = {len(emisoras_fm_final)} total (filtradas {total_filtradas_fm} > 108MHz)")
    log_message(f"📺 TV: {len(emisoras_spectra_tv)} de SPECTRA + {len(emisoras_manuales_tv)} manuales = {len(emisoras_tv_final)} total")
    log_message(f"🔊 AM: {len(emisoras_spectra_am)} de SPECTRA + {len(emisoras_manuales_am)} manuales = {len(emisoras_am_final)} total")

def fusionar_emisoras(emisoras_spectra, emisoras_manuales, callback_log=None):
    """
    Fusiona emisoras de SPECTRA con emisoras manuales
    Prioriza SPECTRA pero mantiene manuales que no tengan conflictos de frecuencia
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    # Crear diccionario de frecuencias de SPECTRA para verificación rápida
    frecuencias_spectra = {emisora['frecuencia'] for emisora in emisoras_spectra}
    
    # Agregar manuales que no entren en conflicto con SPECTRA
    emisoras_manuales_preservadas = []
    for emisora_manual in emisoras_manuales:
        if emisora_manual['frecuencia'] not in frecuencias_spectra:
            emisoras_manuales_preservadas.append(emisora_manual)
        else:
            log_message(f"  ⚠️ Emisora manual con frecuencia {emisora_manual['frecuencia']} MHz reemplazada por datos SPECTRA")
    
    # Combinar SPECTRA + manuales preservadas
    return emisoras_spectra + emisoras_manuales_preservadas

def mapear_servicio_a_tipo(servicio):
    """
    Mapea el servicio de Spectra al tipo en config
    """
    mapeo = {
        'FM - Frecuencia Modulada': 'FM',
        'TV - Televisión Abierta': 'TV',
        'AM - Amplitud Modulada': 'AM'
    }
    
    return mapeo.get(servicio, '')

def formatear_nombre_emisora(red, estado_estacion, estado_solicitud):
    """
    Formatea el nombre de la emisora según los estados
    """
    nombre = red.strip()
    
    # Verificar si está cancelada
    if estado_estacion == 'Cancelado' and estado_solicitud == 'Cancelado':
        # Solo agregar (NO AUTORIZADA) si no la tiene ya
        if '(NO AUTORIZADA)' not in nombre and '(NO AUTORIZADO)' not in nombre:
            nombre += ' (NO AUTORIZADA)'
    
    return nombre

def eliminar_duplicados_por_frecuencia(emisoras):
    """
    Elimina emisoras duplicadas por frecuencia, manteniendo la última encontrada
    """
    emisoras_unicas = {}
    
    for emisora in emisoras:
        frecuencia = emisora['frecuencia']
        emisoras_unicas[frecuencia] = emisora
    
    return list(emisoras_unicas.values())

# Función principal para usar desde la GUI
def actualizar_config_desde_gui(callback_log=None, callback_progress=None):
    """
    Función principal para actualizar configuración desde la GUI
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
        log_message("🎯 INICIANDO ACTUALIZACIÓN DE CONFIGURACIÓN")
        
        # Configuración de rutas
        config_path = "config.json"
        spectra_folder = "SPECTRA_Filtrado"
        
        # Verificar que existe el archivo config.json
        if not os.path.exists(config_path):
            error_msg = f"❌ No se encontró el archivo {config_path}"
            log_message(error_msg)
            return False
        
        # Verificar que existe la carpeta SPECTRA
        if not os.path.exists(spectra_folder):
            error_msg = f"❌ No se encontró la carpeta {spectra_folder}"
            log_message(error_msg)
            return False
        
        # Hacer una copia de seguridad del config original
        backup_path = config_path.replace('.json', '_backup.json')
        shutil.copy2(config_path, backup_path)
        log_message(f"💾 Copia de seguridad creada: {backup_path}")
        
        update_progress(5)
        
        # Actualizar con la nueva función que preserva emisoras manuales
        log_message("\n🔄 ACTUALIZANDO CONFIG PRESERVANDO EMISORAS MANUALES")
        config_actualizado = actualizar_config_desde_spectra(
            config_path, 
            spectra_folder, 
            callback_log, 
            callback_progress
        )
        
        if config_actualizado:
            log_message("✅ ACTUALIZACIÓN DE CONFIGURACIÓN COMPLETADA EXITOSAMENTE")
            return True
        else:
            log_message("❌ FALLÓ LA ACTUALIZACIÓN DE CONFIGURACIÓN")
            return False
        
    except Exception as e:
        error_msg = f"❌ ERROR en actualización de configuración: {str(e)}"
        log_message(error_msg)
        return False

# Ejemplo de uso
if __name__ == "__main__":
    # Para uso directo del script
    actualizar_config_desde_gui()