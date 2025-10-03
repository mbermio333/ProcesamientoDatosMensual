import json
import os
from pathlib import Path
from datetime import datetime

def normalizar_frecuencias_tv(archivo_json):
    """
    Normaliza las frecuencias de TV en un archivo JSON y REEMPLAZA el archivo original:
    - Mantiene frecuencias decimales (que contienen ".")
    - A frecuencias enteras les resta 1.75
    - Para FRECUENCIAS duplicadas, prioriza registros "Activos"
    """
    # Cargar datos
    with open(archivo_json, 'r', encoding='utf-8') as f:
        datos = json.load(f)
    
    print(f"Procesando {len(datos)} registros en {archivo_json}")
    
    # Filtrar solo registros de TV
    registros_tv = [registro for registro in datos if registro.get('SERVICIO') == 'TV - Televisión Abierta']
    otros_registros = [registro for registro in datos if registro.get('SERVICIO') != 'TV - Televisión Abierta']
    
    print(f"Encontrados {len(registros_tv)} registros de TV para normalizar")
    
    # Procesar cada registro de TV individualmente
    for registro in registros_tv:
        frecuencia = registro.get('FRECUENCIA')
        if frecuencia is not None:
            # Convertir a string para verificar si es decimal
            frecuencia_str = str(frecuencia)
            
            if '.' in frecuencia_str:
                # Mantener frecuencia decimal sin cambios
                print(f"Manteniendo frecuencia decimal: {frecuencia} para {registro.get('RED')}")
            else:
                # Restar 1.75 a frecuencia entera
                try:
                    frecuencia_float = float(frecuencia)
                    nueva_frecuencia = frecuencia_float - 1.75
                    print(f"Ajustando frecuencia entera: {frecuencia} -> {nueva_frecuencia} para {registro.get('RED')}")
                    registro['FRECUENCIA'] = nueva_frecuencia
                except (ValueError, TypeError):
                    print(f"Error procesando frecuencia: {frecuencia} para {registro.get('RED')}")
                    continue
    
    # Ahora eliminar duplicados por FRECUENCIA (después de la normalización)
    registros_tv_sin_duplicados = eliminar_duplicados_por_frecuencia(registros_tv)
    
    print(f"Registros TV después de eliminar duplicados: {len(registros_tv_sin_duplicados)}")
    
    # Combinar registros procesados de TV con los otros registros
    datos_finales = otros_registros + registros_tv_sin_duplicados
    
    # GUARDAR SOBREESCRIBIENDO EL ARCHIVO ORIGINAL
    with open(archivo_json, 'w', encoding='utf-8') as f:
        json.dump(datos_finales, f, ensure_ascii=False, indent=2)
    
    print(f"Procesamiento completado. Archivo ORIGINAL actualizado: {archivo_json}")
    
    return datos_finales, archivo_json

def eliminar_duplicados_por_frecuencia(registros_tv):
    """
    Elimina registros duplicados por FRECUENCIA, priorizando los que están "Activos"
    """
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
            print(f"\nEncontrados {len(registros)} registros para frecuencia: {clave} MHz")
            
            # Mostrar todos los registros encontrados
            for i, r in enumerate(registros, 1):
                estado_est = r.get('ESTADO_ESTACION', 'N/A')
                estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                red = r.get('RED', 'N/A')
                print(f"  {i}. {red} - Estados: {estado_est}/{estado_sol}")
            
            # Priorizar registros activos
            registros_activos = [r for r in registros if 
                               r.get('ESTADO_ESTACION') == 'Activo' or 
                               r.get('ESTADO_SOLICITUD') == 'Activo']
            
            if registros_activos:
                # Entre activos, elegir el de mayor vigencia
                registro_elegido = max(registros_activos, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                print(f"✓ Elegido registro ACTIVO: {registro_elegido.get('RED')}")
                
                # Mostrar los descartados
                for r in registros:
                    if r != registro_elegido:
                        estado_est = r.get('ESTADO_ESTACION', 'N/A')
                        estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                        print(f"  ✗ Descartado: {r.get('RED')} - Estados: {estado_est}/{estado_sol}")
            else:
                # Si no hay activos, elegir el de mayor vigencia entre todos
                registro_elegido = max(registros, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                print(f"✓ Elegido registro por VIGENCIA: {registro_elegido.get('RED')}")
    
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

def procesar_todos_los_archivos_en_carpeta(carpeta):
    """
    Procesa todos los archivos JSON en una carpeta y REEMPLAZA los archivos originales
    """
    carpeta_path = Path(carpeta)
    archivos_json = list(carpeta_path.glob('*.json'))
    
    print(f"Encontrados {len(archivos_json)} archivos para procesar")
    
    resultados = {}
    
    for archivo in archivos_json:
        print(f"\n{'='*50}")
        print(f"PROCESANDO: {archivo.name}")
        print(f"{'='*50}")
        
        try:
            datos_normalizados, archivo_salida = normalizar_frecuencias_tv(str(archivo))
            resultados[archivo.name] = {
                'archivo_procesado': str(archivo),
                'total_registros': len(datos_normalizados)
            }
        except Exception as e:
            print(f"Error procesando {archivo.name}: {str(e)}")
            resultados[archivo.name] = {'error': str(e)}
    
    print(f"\n{'='*60}")
    print(f"PROCESAMIENTO COMPLETADO")
    print(f"{'='*60}")
    print(f"Total de archivos procesados: {len(resultados)}")
    print(f"Todos los archivos originales han sido actualizados")
    
    return resultados

# Ejemplo de uso
if __name__ == "__main__":
    # Procesar todos los archivos en una carpeta
    carpeta_resultados = "SPECTRA_filtrado"
    resultados = procesar_todos_los_archivos_en_carpeta(carpeta_resultados)