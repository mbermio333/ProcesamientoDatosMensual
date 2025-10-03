import openpyxl
import json
from datetime import datetime
import os
from pathlib import Path

def filtrar_datos_arcotel(archivo_excel, archivo_salida, filtros_personalizados=None):
    """
    Filtra datos de la base de datos ARCOTEL según criterios específicos y guarda en JSON
    
    Args:
        archivo_excel (str): Ruta del archivo Excel de entrada
        archivo_salida (str): Ruta completa del archivo JSON de salida
        filtros_personalizados (dict): Diccionario con filtros personalizados
    """
    try:
        # Crear directorio si no existe
        directorio_salida = os.path.dirname(archivo_salida)
        if directorio_salida and not os.path.exists(directorio_salida):
            os.makedirs(directorio_salida)
            print(f"Directorio creado: {directorio_salida}")
        
        # Cargar el archivo Excel
        workbook = openpyxl.load_workbook(archivo_excel)
        sheet = workbook['descarga']
        
        # Usar filtros personalizados o por defecto
        if filtros_personalizados is None:
            filtros = {
                'PROVINCIA_A': 'ZAMORA CHINCHIPE',
                'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Television Abierta'],
                'AREAS_OP_MR': 'ZAMORA'
            }
        else:
            filtros = filtros_personalizados
        
        # Mapeo de columnas a extraer
        columnas_extraer = {
            'SERVICIO': 'B',
            'NOMBRES': 'C',
            'RED': 'E', 
            'FRECUENCIA': 'F',
            'ESTADO_ESTACION': 'AI',
            'ESTACION': 'AR',
            'ESTADO_SOLICITUD': 'BC',
            'SUSCRIPCION': 'AJ',
            'VIGENCIA': 'AK'
        }
        
        # Lista para almacenar los resultados iniciales
        resultados_iniciales = []
        
        # Contadores para estadísticas
        total_filas = sheet.max_row - 1
        filas_procesadas = 0
        
        ciudad = filtros.get('AREAS_OP_MR', 'DESCONOCIDA')
        print(f"Procesando {total_filas} filas para ciudad: {ciudad}...")
        
        # Iterar sobre las filas
        for row in range(2, sheet.max_row + 1):
            filas_procesadas += 1
            if filas_procesadas % 1000 == 0:
                print(f"Procesadas {filas_procesadas}/{total_filas} filas...")
            
            # Obtener valores de las columnas de filtro
            provincia = sheet[f'A{row}'].value
            servicio = sheet[f'B{row}'].value
            areas_op = sheet[f'I{row}'].value
            
            # Aplicar filtros
            cumple_filtros = True
            
            # Filtro PROVINCIA_A (exacto)
            if 'PROVINCIA_A' in filtros and provincia != filtros['PROVINCIA_A']:
                cumple_filtros = False
            
            # FILTRO SERVICIO - acepta múltiples servicios
            if 'SERVICIOS' in filtros and servicio not in filtros['SERVICIOS']:
                cumple_filtros = False
            
            # FILTRO AREAS_OP_MR - busca contenido
            if 'AREAS_OP_MR' in filtros and (areas_op is None or 
               filtros['AREAS_OP_MR'].upper() not in str(areas_op).upper()):
                cumple_filtros = False
            
            if cumple_filtros:
                registro = {}
                for nombre_col, col in columnas_extraer.items():
                    valor = sheet[f'{col}{row}'].value
                    
                    if isinstance(valor, datetime):
                        valor = valor.strftime('%Y-%m-%d %H:%M:%S')
                    elif valor is None:
                        valor = ""
                    
                    registro[nombre_col] = valor
                
                registro['_FILTRO_AREAS_OP_MR'] = areas_op
                registro['_ROW_NUMBER'] = row
                resultados_iniciales.append(registro)
        
        print(f"Filtros básicos aplicados. Encontrados {len(resultados_iniciales)} registros para {ciudad}.")
        
        # Aplicar filtro por vigencia
        resultados_finales = filtrar_por_vigencia_mas_alta(resultados_iniciales)
        
        # Guardar resultados
        with open(archivo_salida, 'w', encoding='utf-8') as json_file:
            json.dump(resultados_finales, json_file, ensure_ascii=False, indent=2)
        
        print(f"Proceso completado para {ciudad}. Registros únicos: {len(resultados_finales)}")
        print(f"Archivo guardado en: {archivo_salida}")
        
        # Mostrar estadísticas
        if resultados_finales:
            mostrar_estadisticas(resultados_iniciales, resultados_finales, ciudad)
        
        return resultados_finales
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {archivo_excel}")
        return []
    except Exception as e:
        print(f"Error durante el procesamiento para {filtros.get('AREAS_OP_MR', 'ciudad desconocida')}: {str(e)}")
        return []

def procesar_multiple_ciudades(archivo_excel, directorio_salida, configuraciones_ciudades):
    """
    Procesa múltiples ciudades y guarda los resultados en una misma carpeta
    
    Args:
        archivo_excel (str): Ruta del archivo Excel
        directorio_salida (str): Directorio donde guardar los archivos JSON
        configuraciones_ciudades (list): Lista de diccionarios con configuraciones por ciudad
    """
    print(f"=== INICIANDO PROCESAMIENTO PARA {len(configuraciones_ciudades)} CIUDADES ===")
    print(f"Directorio de salida: {directorio_salida}")
    
    resultados_totales = {}
    
    for config in configuraciones_ciudades:
        ciudad = config['AREAS_OP_MR']
        nombre_archivo = f"resultados_{ciudad.lower().replace(' ', '_')}.json"
        archivo_salida = os.path.join(directorio_salida, nombre_archivo)
        
        print(f"\n--- Procesando: {ciudad} ---")
        
        resultados = filtrar_datos_arcotel(archivo_excel, archivo_salida, config)
        resultados_totales[ciudad] = resultados
    
    # Generar resumen general
    generar_resumen_general(directorio_salida, resultados_totales)
    
    return resultados_totales

def generar_resumen_general(directorio_salida, resultados_totales):
    """Genera un archivo de resumen con estadísticas de todas las ciudades"""
    resumen = {
        'fecha_procesamiento': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_ciudades': len(resultados_totales),
        'estadisticas_por_ciudad': {}
    }
    
    total_registros = 0
    
    for ciudad, resultados in resultados_totales.items():
        resumen['estadisticas_por_ciudad'][ciudad] = {
            'total_registros': len(resultados),
            'servicios': {}
        }
        total_registros += len(resultados)
        
        # Contar por servicio
        for registro in resultados:
            servicio = registro.get('RED', 'Sin servicio')
            resumen['estadisticas_por_ciudad'][ciudad]['servicios'][servicio] = \
                resumen['estadisticas_por_ciudad'][ciudad]['servicios'].get(servicio, 0) + 1
    
    resumen['total_registros_general'] = total_registros
    
    # Guardar resumen
    archivo_resumen = os.path.join(directorio_salida, "resumen_general.json")
    with open(archivo_resumen, 'w', encoding='utf-8') as f:
        json.dump(resumen, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== RESUMEN GENERAL ===")
    print(f"Total de ciudades procesadas: {len(resultados_totales)}")
    print(f"Total de registros únicos: {total_registros}")
    print(f"Resumen guardado en: {archivo_resumen}")
    
    for ciudad, stats in resumen['estadisticas_por_ciudad'].items():
        print(f"\n{ciudad}:")
        print(f"  - Registros: {stats['total_registros']}")
        for servicio, cantidad in stats['servicios'].items():
            print(f"  - {servicio}: {cantidad}")

# Las funciones filtrar_por_vigencia_mas_alta y mostrar_estadisticas se mantienen igual
def filtrar_por_vigencia_mas_alta(registros):
    """Filtra registros por frecuencia, manteniendo solo el registro con la vigencia más alta"""
    registros_por_frecuencia = {}
    
    for registro in registros:
        frecuencia = registro.get('FRECUENCIA', '')
        vigencia_str = registro.get('VIGENCIA', '')
        
        vigencia_dt = None
        if vigencia_str:
            try:
                if isinstance(vigencia_str, datetime):
                    vigencia_dt = vigencia_str
                else:
                    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
                        try:
                            vigencia_dt = datetime.strptime(str(vigencia_str), fmt)
                            break
                        except ValueError:
                            continue
            except (ValueError, TypeError):
                vigencia_dt = None
        
        if not frecuencia:
            frecuencia = f"sin_frecuencia_{id(registro)}"
        
        if frecuencia not in registros_por_frecuencia:
            registros_por_frecuencia[frecuencia] = registro
            registro['_VIGENCIA_DT'] = vigencia_dt
        else:
            registro_existente = registros_por_frecuencia[frecuencia]
            vigencia_existente = registro_existente.get('_VIGENCIA_DT')
            
            if vigencia_dt and vigencia_existente:
                if vigencia_dt > vigencia_existente:
                    registros_por_frecuencia[frecuencia] = registro
                    registro['_VIGENCIA_DT'] = vigencia_dt
            elif vigencia_dt and not vigencia_existente:
                registros_por_frecuencia[frecuencia] = registro
                registro['_VIGENCIA_DT'] = vigencia_dt
    
    resultados_finales = []
    for frecuencia, registro in registros_por_frecuencia.items():
        registro_final = {k: v for k, v in registro.items() if not k.startswith('_')}
        resultados_finales.append(registro_final)
    
    return resultados_finales

def mostrar_estadisticas(registros_iniciales, registros_finales, ciudad):
    """Muestra estadísticas comparativas"""
    print(f"\n=== ESTADÍSTICAS PARA {ciudad} ===")
    print(f"Registros iniciales: {len(registros_iniciales)}")
    print(f"Registros finales: {len(registros_finales)}")
    print(f"Registros eliminados: {len(registros_iniciales) - len(registros_finales)}")

if __name__ == "__main__":
    # Configuración principal
    archivo_excel = "9. SPECTRA_10sep2025.xlsx"
    directorio_base = "SPECTRA_Filtrado"  # Carpeta donde se guardarán todos los archivos
    
    # Definir las ciudades a procesar
    configuraciones_ciudades = [
        {
            'PROVINCIA_A': 'ZAMORA CHINCHIPE',
            'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Televisión Abierta'],
            'AREAS_OP_MR': 'ZAMORA'
        },
        {
            'PROVINCIA_A': 'LOJA',
            'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Televisión Abierta'],
            'AREAS_OP_MR': 'LOJA'
        },
        {
            'PROVINCIA_A': 'CAÑAR', 
            'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Televisión Abierta'],
            'AREAS_OP_MR': 'TAMBO'
        },
        {
            'PROVINCIA_A': 'MORONA SANTIAGO',
            'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Televisión Abierta'],
            'AREAS_OP_MR': 'MORONA'
        },
        {
            'PROVINCIA_A': 'EL ORO',
            'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Televisión Abierta'],
            'AREAS_OP_MR': 'MACHALA'
        },
        {
            'PROVINCIA_A': 'AZUAY',
            'SERVICIOS': ['FM - Frecuencia Modulada','AM - Amplitud Modulada','TV - Televisión Abierta'],
            'AREAS_OP_MR': 'CUENCA'
        }


        # Puedes agregar más ciudades aquí
    ]
    
    # Opción 1: Procesar múltiples ciudades
    resultados = procesar_multiple_ciudades(archivo_excel, directorio_base, configuraciones_ciudades)
    
    # Opción 2: Procesar una sola ciudad (manteniendo compatibilidad)
    # archivo_salida = os.path.join(directorio_base, "resultados_zamora.json")
    # resultados = filtrar_datos_arcotel(archivo_excel, archivo_salida)