import json
import os
from pathlib import Path
from datetime import datetime

class DecimalEncoder(json.JSONEncoder):
    """Encoder personalizado para formatear floats según el servicio"""
    def encode(self, obj):
        if isinstance(obj, float):
            # Para JSON general, usar 2 decimales por defecto
            return format(obj, '.2f')
        return super().encode(obj)

def normalizar_frecuencias_tv_fm_am(archivo_json, callback_log=None):
    """
    Normaliza las frecuencias en un archivo JSON y REEMPLAZA el archivo original:
    
    Para TV:
    - Mantiene frecuencias decimales (que contienen ".")
    - A frecuencias enteras les resta 1.75
    - FORMATO: 2 decimales
    
    Para FM:
    - Trunca frecuencias con más de 1 decimal a solo 1 decimal
    - FORMATO: 1 decimal
    
    Para AM:
    - Convierte de KHz a MHz (dividiendo entre 1000)
    - FORMATO: 3 decimales (para mayor precisión en AM)
    
    Para ambos:
    - Elimina duplicados por FRECUENCIA, priorizando registros "Activos"
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
    
    # Procesar cada registro según su servicio
    registros_modificados = 0
    for registro in datos:
        servicio = registro.get('SERVICIO')
        frecuencia = registro.get('FRECUENCIA')
        
        if frecuencia is None:
            continue
            
        if servicio == 'TV - Televisión Abierta':
            # Procesar TV - 2 decimales
            if procesar_frecuencia_tv(registro, callback_log):
                registros_modificados += 1
        elif servicio == 'FM - Frecuencia Modulada':
            # Procesar FM - 1 decimal
            if procesar_frecuencia_fm(registro, callback_log):
                registros_modificados += 1
        elif servicio == 'AM - Amplitud Modulada':
            # Procesar AM - convertir KHz a MHz
            if procesar_frecuencia_am(registro, callback_log):
                registros_modificados += 1
        # Para otros servicios, no hacer cambios
    
    log_message(f"🔄 Se modificaron {registros_modificados} registros de frecuencia")
    
    # Eliminar duplicados por FRECUENCIA (después de la normalización)
    datos_sin_duplicados = eliminar_duplicados_por_frecuencia(datos, callback_log)
    
    log_message(f"✅ Registros después de eliminar duplicados: {len(datos_sin_duplicados)}")
    
    # GUARDAR SOBREESCRIBIENDO EL ARCHIVO ORIGINAL con formato forzado
    with open(archivo_json, 'w', encoding='utf-8') as f:
        # Usar el encoder personalizado
        json.dump(datos_sin_duplicados, f, ensure_ascii=False, indent=2, cls=DecimalEncoder)
    
    log_message(f"💾 Procesamiento completado. Archivo ORIGINAL actualizado: {os.path.basename(archivo_json)}")
    
    return datos_sin_duplicados, archivo_json

def procesar_frecuencia_tv(registro, callback_log=None):
    """Procesa frecuencia de TV según las reglas específices - 2 decimales"""
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    frecuencia = registro.get('FRECUENCIA')
    if frecuencia is None:
        return False
    
    modificado = False
    frecuencia_original = frecuencia
    
    try:
        frecuencia_float = float(frecuencia)
        
        # Convertir a string para verificar si es decimal
        frecuencia_str = str(frecuencia)
        
        if '.' in frecuencia_str:
            # Mantener frecuencia decimal pero forzar a 2 decimales
            nueva_frecuencia = round(frecuencia_float, 2)
            if abs(frecuencia_float - nueva_frecuencia) > 0.0001:
                registro['FRECUENCIA'] = nueva_frecuencia
                log_message(f"📺 TV - Formateando a 2 decimales: {frecuencia_original} -> {nueva_frecuencia} para {registro.get('RED')}")
                modificado = True
            else:
                # Ya está normalizada, pero forzar el formato a 2 decimales
                registro['FRECUENCIA'] = nueva_frecuencia
                modificado = True
        else:
            # Restar 1.75 a frecuencia entera y forzar a 2 decimales
            nueva_frecuencia = round(frecuencia_float - 1.75, 2)
            log_message(f"📺 TV - Ajustando frecuencia entera: {frecuencia_original} -> {nueva_frecuencia} para {registro.get('RED')}")
            registro['FRECUENCIA'] = nueva_frecuencia
            modificado = True
            
    except (ValueError, TypeError):
        log_message(f"❌ TV - Error procesando frecuencia: {frecuencia} para {registro.get('RED')}")
    
    return modificado

def procesar_frecuencia_fm(registro, callback_log=None):
    """Procesa frecuencia de FM: trunca a 1 decimal"""
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    frecuencia = registro.get('FRECUENCIA')
    if frecuencia is None:
        return False
    
    try:
        frecuencia_float = float(frecuencia)
        frecuencia_original = frecuencia
        
        # Forzar a exactamente 1 decimal
        frecuencia_normalizada = round(frecuencia_float, 1)
        
        # Verificar si realmente necesitamos cambiar
        if abs(frecuencia_float - frecuencia_normalizada) > 0.000001:
            log_message(f"📻 FM - Normalizando a 1 decimal: {frecuencia_original} -> {frecuencia_normalizada} para {registro.get('RED')}")
            registro['FRECUENCIA'] = frecuencia_normalizada
            return True
        else:
            # Ya está normalizada, pero forzar el formato a 1 decimal
            registro['FRECUENCIA'] = frecuencia_normalizada
            return True
            
    except (ValueError, TypeError):
        log_message(f"❌ FM - Error procesando frecuencia: {frecuencia} para {registro.get('RED')}")
        return False

def procesar_frecuencia_am(registro, callback_log=None):
    """Procesa frecuencia de AM: convierte de KHz a MHz - 3 decimales"""
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    frecuencia = registro.get('FRECUENCIA')
    if frecuencia is None:
        return False
    
    try:
        frecuencia_float = float(frecuencia)
        frecuencia_original = frecuencia
        
        # Convertir de KHz a MHz (dividir entre 1000)
        frecuencia_mhz = frecuencia_float / 1000.0
        
        # Redondear a 3 decimales para mayor precisión en AM
        frecuencia_normalizada = round(frecuencia_mhz, 3)
        
        log_message(f"📡 AM - Convirtiendo KHz a MHz: {frecuencia_original} KHz -> {frecuencia_normalizada} MHz para {registro.get('RED')}")
        registro['FRECUENCIA'] = frecuencia_normalizada
        return True
            
    except (ValueError, TypeError):
        log_message(f"❌ AM - Error procesando frecuencia: {frecuencia} para {registro.get('RED')}")
        return False

def eliminar_duplicados_por_frecuencia(registros, callback_log=None):
    """
    Elimina registros duplicados por FRECUENCIA, priorizando los que están "Activos"
    """
    
    def log_message(message):
        if callback_log:
            callback_log(message)
        else:
            print(message)
    
    registros_por_frecuencia = {}
    
    for registro in registros:
        frecuencia = registro.get('FRECUENCIA')
        servicio = registro.get('SERVICIO', 'Desconocido')
        
        if frecuencia is None:
            # Si no tiene frecuencia, mantenerlo
            continue
            
        # Normalizar frecuencia para la clave según el servicio
        try:
            if servicio == 'FM - Frecuencia Modulada':
                frecuencia_clave = round(float(frecuencia), 1)  # 1 decimal para FM
            elif servicio == 'TV - Televisión Abierta':
                frecuencia_clave = round(float(frecuencia), 2)  # 2 decimales para TV
            elif servicio == 'AM - Amplitud Modulada':
                frecuencia_clave = round(float(frecuencia), 3)  # 3 decimales para AM
            else:
                frecuencia_clave = float(frecuencia)  # Sin formato específico para otros
        except (ValueError, TypeError):
            frecuencia_clave = frecuencia
        
        # Crear clave única por SERVICIO y FRECUENCIA normalizada
        clave = f"{servicio}_{frecuencia_clave}"
        
        if clave not in registros_por_frecuencia:
            registros_por_frecuencia[clave] = []
        registros_por_frecuencia[clave].append(registro)
    
    # Para cada grupo de registros con misma FRECUENCIA y SERVICIO, elegir el mejor
    registros_unicos = []
    
    for clave, registros_grupo in registros_por_frecuencia.items():
        servicio = clave.split('_')[0]
        frecuencia = '_'.join(clave.split('_')[1:])
        
        if len(registros_grupo) == 1:
            # Solo un registro, mantenerlo
            registros_unicos.append(registros_grupo[0])
        else:
            # Múltiples registros con misma frecuencia, elegir el mejor
            log_message(f"\n🔍 Encontrados {len(registros_grupo)} registros para {servicio} frecuencia: {frecuencia} MHz")
            
            # Mostrar todos los registros encontrados
            for i, r in enumerate(registros_grupo, 1):
                estado_est = r.get('ESTADO_ESTACION', 'N/A')
                estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                red = r.get('RED', 'N/A')
                log_message(f"  {i}. {red} - Estados: {estado_est}/{estado_sol}")
            
            # Priorizar registros activos
            registros_activos = [r for r in registros_grupo if 
                               r.get('ESTADO_ESTACION') == 'Activo' or 
                               r.get('ESTADO_SOLICITUD') == 'Activo']
            
            if registros_activos:
                # Entre activos, elegir el de mayor vigencia
                registro_elegido = max(registros_activos, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                log_message(f"✅ Elegido registro ACTIVO: {registro_elegido.get('RED')}")
                
                # Mostrar los descartados
                for r in registros_grupo:
                    if r != registro_elegido:
                        estado_est = r.get('ESTADO_ESTACION', 'N/A')
                        estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                        log_message(f"  ❌ Descartado: {r.get('RED')} - Estados: {estado_est}/{estado_sol}")
            else:
                # Si no hay activos, elegir el de mayor vigencia entre todos
                registro_elegido = max(registros_grupo, 
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
            datos_normalizados, archivo_salida = normalizar_frecuencias_tv_fm_am(str(archivo), callback_log)
            resultados[archivo.name] = {
                'archivo_procesado': str(archivo),
                'total_registros': len(datos_normalizados),
                'estado': 'Completado'
            }
            
            # Actualizar progreso
            if callback_progress:
                progreso = 10 + (i / len(archivos_json)) * 80  # 10% a 90%
                update_progress(int(progreso))
                
        except Exception as e:
            error_msg = f"❌ Error procesando {archivo.name}: {str(e)}"
            log_message(error_msg)
            resultados[archivo.name] = {'error': str(e), 'estado': 'Error'}
    
    log_message(f"\n{'='*60}")
    log_message(f"✅ PROCESAMIENTO COMPLETADO")
    log_message(f"{'='*60}")
    log_message(f"📊 Total de archivos procesados: {len(resultados)}")
    log_message(f"💾 Todos los archivos originales han sido actualizados")
    
    return resultados

# Función principal para usar desde la GUI
def procesar_normalizacion_tv_fm_am_desde_gui(callback_log=None, callback_progress=None):
    """
    Función principal para normalizar TV, FM y AM desde la GUI
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
        log_message("🎯 INICIANDO NORMALIZACIÓN DE FRECUENCIAS TV, FM Y AM")
        log_message("📺 TV: 2 decimales | 📻 FM: 1 decimal | 📡 AM: KHz a MHz (3 decimales)")
        
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
        
        log_message("✅ NORMALIZACIÓN TV, FM Y AM COMPLETADA EXITOSAMENTE")
        update_progress(100)
        
        return True
        
    except Exception as e:
        error_msg = f"❌ ERROR en normalización TV, FM y AM: {str(e)}"
        log_message(error_msg)
        return False

# Ejemplo de uso
if __name__ == "__main__":
    # Para uso directo del script
    procesar_normalizacion_tv_fm_am_desde_gui()