import json
import os
from pathlib import Path
from datetime import datetime

def normalizar_frecuencias_tv(archivo_json):
    """
    Normaliza las frecuencias de TV en un archivo JSON:
    - Mantiene frecuencias decimales (que contienen ".")
    - A frecuencias enteras les resta 1.75
    - Para REDES duplicadas, prioriza registros "Activos"
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
                registro['FRECUENCIA_ORIGINAL'] = frecuencia  # Opcional: guardar original
            else:
                # Restar 1.75 a frecuencia entera
                try:
                    frecuencia_float = float(frecuencia)
                    nueva_frecuencia = frecuencia_float - 1.75
                    print(f"Ajustando frecuencia entera: {frecuencia} -> {nueva_frecuencia} para {registro.get('RED')}")
                    registro['FRECUENCIA_ORIGINAL'] = frecuencia  # Guardar original
                    registro['FRECUENCIA'] = nueva_frecuencia
                except (ValueError, TypeError):
                    print(f"Error procesando frecuencia: {frecuencia} para {registro.get('RED')}")
                    continue
    
    # Ahora eliminar duplicados por RED y FRECUENCIA (después de la normalización)
    registros_tv_sin_duplicados = eliminar_duplicados_por_red(registros_tv)
    
    print(f"Registros TV después de eliminar duplicados: {len(registros_tv_sin_duplicados)}")
    
    # Combinar registros procesados de TV con los otros registros
    datos_finales = otros_registros + registros_tv_sin_duplicados
    
    # Guardar el archivo normalizado
    archivo_salida = archivo_json.replace('.json', '_normalizado.json')
    with open(archivo_salida, 'w', encoding='utf-8') as f:
        json.dump(datos_finales, f, ensure_ascii=False, indent=2)
    
    print(f"Procesamiento completado. Archivo guardado en: {archivo_salida}")
    
    return datos_finales, archivo_salida

def eliminar_duplicados_por_red(registros_tv):
    """
    Elimina registros duplicados por RED, priorizando los que están "Activos"
    """
    registros_por_red = {}
    
    for registro in registros_tv:
        red = registro.get('RED', 'Sin nombre')
        frecuencia = registro.get('FRECUENCIA')
        
        # Crear clave única por RED y FRECUENCIA
        clave = f"{red}_{frecuencia}"
        
        if clave not in registros_por_red:
            registros_por_red[clave] = []
        registros_por_red[clave].append(registro)
    
    # Para cada grupo de registros con misma RED y FRECUENCIA, elegir el mejor
    registros_unicos = []
    
    for clave, registros in registros_por_red.items():
        if len(registros) == 1:
            # Solo un registro, mantenerlo
            registros_unicos.append(registros[0])
        else:
            # Múltiples registros, elegir el mejor
            print(f"\nEncontrados {len(registros)} registros para: {clave}")
            
            # Priorizar registros activos
            registros_activos = [r for r in registros if 
                               r.get('ESTADO_ESTACION') == 'Activo' or 
                               r.get('ESTADO_SOLICITUD') == 'Activo']
            
            if registros_activos:
                # Entre activos, elegir el de mayor vigencia
                registro_elegido = max(registros_activos, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                print(f"✓ Elegido registro ACTIVO: {registro_elegido.get('NOMBRES')}")
                
                # Mostrar los descartados
                for r in registros:
                    if r != registro_elegido:
                        estado_est = r.get('ESTADO_ESTACION', 'N/A')
                        estado_sol = r.get('ESTADO_SOLICITUD', 'N/A')
                        print(f"  ✗ Descartado: {r.get('NOMBRES')} - Estados: {estado_est}/{estado_sol}")
            else:
                # Si no hay activos, elegir el de mayor vigencia entre todos
                registro_elegido = max(registros, 
                                     key=lambda x: convertir_vigencia_a_fecha(x.get('VIGENCIA', '')))
                registros_unicos.append(registro_elegido)
                print(f"✓ Elegido registro por VIGENCIA: {registro_elegido.get('NOMBRES')}")
    
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
    Procesa todos los archivos JSON en una carpeta
    """
    carpeta_path = Path(carpeta)
    archivos_json = list(carpeta_path.glob('*.json'))
    
    # Excluir archivos que ya están normalizados
    archivos_a_procesar = [f for f in archivos_json 
                          if not f.name.endswith('_normalizado.json')]
    
    print(f"Encontrados {len(archivos_a_procesar)} archivos para procesar")
    
    resultados = {}
    
    for archivo in archivos_a_procesar:
        print(f"\n{'='*50}")
        print(f"PROCESANDO: {archivo.name}")
        print(f"{'='*50}")
        
        try:
            datos_normalizados, archivo_salida = normalizar_frecuencias_tv(str(archivo))
            resultados[archivo.name] = {
                'archivo_original': str(archivo),
                'archivo_normalizado': archivo_salida,
                'total_registros': len(datos_normalizados)
            }
        except Exception as e:
            print(f"Error procesando {archivo.name}: {str(e)}")
            resultados[archivo.name] = {'error': str(e)}
    
    print(f"\n{'='*60}")
    print(f"PROCESAMIENTO COMPLETADO")
    print(f"{'='*60}")
    print(f"Total de archivos procesados: {len(resultados)}")
    
    return resultados

# Función simplificada para un solo archivo
def normalizacion_rapida(archivo_json):
    """
    Versión simplificada para normalizar un archivo específico
    """
    return normalizar_frecuencias_tv(archivo_json)

# Ejemplo de uso
if __name__ == "__main__":
    # Opción 1: Procesar un archivo individual
    # archivo_individual = "zamora_spectra.json"
    # datos_normalizados, archivo_salida = normalizacion_rapida(archivo_individual)
    
    # Opción 2: Procesar todos los archivos en una carpeta (RECOMENDADO)
    carpeta_resultados = "SPECTRA_filtrado"  # La carpeta donde están tus JSON
    resultados = procesar_todos_los_archivos_en_carpeta(carpeta_resultados)