import openpyxl
import json
from datetime import datetime

def filtrar_datos_arcotel(archivo_excel, archivo_json):
    """
    Filtra datos de la base de datos ARCOTEL según criterios específicos y guarda en JSON
    """
    try:
        # Cargar el archivo Excel
        workbook = openpyxl.load_workbook(archivo_excel)
        sheet = workbook['descarga']  # Usar la hoja 'descarga'
        
        # Definir los filtros
        filtros = {
            'PROVINCIA_A': 'ZAMORA CHINCHIPE',
            'SERVICIOS': ['FM - Frecuencia Modulada', 'TV - Televisión Abierta'],
            'AREAS_OP_MR': 'ZAMORA'
        }
        
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
        
        # Obtener índices de las columnas de filtro
        col_provincia = 'A'  # PROVINCIA_A
        col_servicio = 'B'   # SERVICIO
        col_areas_op = 'I'   # AREAS_OP_MR
        
        # Lista para almacenar los resultados iniciales
        resultados_iniciales = []
        
        # Contadores para estadísticas
        total_filas = sheet.max_row - 1
        filas_procesadas = 0
        
        print(f"Procesando {total_filas} filas...")
        
        # Iterar sobre las filas (empezando desde la fila 2 para saltar los encabezados)
        for row in range(2, sheet.max_row + 1):
            filas_procesadas += 1
            if filas_procesadas % 1000 == 0:
                print(f"Procesadas {filas_procesadas}/{total_filas} filas...")
            
            # Obtener valores de las columnas de filtro
            provincia = sheet[f'{col_provincia}{row}'].value
            servicio = sheet[f'{col_servicio}{row}'].value
            areas_op = sheet[f'{col_areas_op}{row}'].value
            
            # Aplicar filtros
            cumple_filtros = True
            
            # Filtro PROVINCIA_A (exacto)
            if provincia != filtros['PROVINCIA_A']:
                cumple_filtros = False
            
            # FILTRO SERVICIO - acepta múltiples servicios
            if servicio not in filtros['SERVICIOS']:
                cumple_filtros = False
            
            # FILTRO AREAS_OP_MR - busca contenido
            if areas_op is None or filtros['AREAS_OP_MR'].upper() not in str(areas_op).upper():
                cumple_filtros = False
            
            if cumple_filtros:
                # Crear diccionario con los datos a extraer
                registro = {}
                
                for nombre_col, col in columnas_extraer.items():
                    valor = sheet[f'{col}{row}'].value
                    
                    # Convertir fechas a string si es necesario
                    if isinstance(valor, datetime):
                        valor = valor.strftime('%Y-%m-%d %H:%M:%S')
                    elif valor is None:
                        valor = ""
                    
                    registro[nombre_col] = valor
                
                # Agregar información de filtro para referencia
                registro['_FILTRO_AREAS_OP_MR'] = areas_op
                registro['_ROW_NUMBER'] = row  # Para debugging
                resultados_iniciales.append(registro)
        
        print(f"\nFiltros básicos aplicados. Encontrados {len(resultados_iniciales)} registros.")
        
        # APLICAR FILTRO POR FRECUENCIA Y VIGENCIA
        resultados_finales = filtrar_por_vigencia_mas_alta(resultados_iniciales)
        
        # Guardar resultados en archivo JSON
        with open(archivo_json, 'w', encoding='utf-8') as json_file:
            json.dump(resultados_finales, json_file, ensure_ascii=False, indent=2)
        
        print(f"Proceso completado. Se encontraron {len(resultados_finales)} registros únicos por frecuencia.")
        print(f"Resultados guardados en: {archivo_json}")
        
        # Mostrar estadísticas
        if resultados_finales:
            mostrar_estadisticas(resultados_iniciales, resultados_finales)
        
        return resultados_finales
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {archivo_excel}")
        return []
    except Exception as e:
        print(f"Error durante el procesamiento: {str(e)}")
        return []

def filtrar_por_vigencia_mas_alta(registros):
    """
    Filtra registros por frecuencia, manteniendo solo el registro con la vigencia más alta por cada frecuencia
    """
    # Agrupar registros por frecuencia
    registros_por_frecuencia = {}
    
    for registro in registros:
        frecuencia = registro.get('FRECUENCIA', '')
        vigencia_str = registro.get('VIGENCIA', '')
        
        # Convertir vigencia a datetime para comparación
        vigencia_dt = None
        if vigencia_str:
            try:
                # Intentar diferentes formatos de fecha
                if isinstance(vigencia_str, datetime):
                    vigencia_dt = vigencia_str
                else:
                    # Probar diferentes formatos de fecha
                    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
                        try:
                            vigencia_dt = datetime.strptime(str(vigencia_str), fmt)
                            break
                        except ValueError:
                            continue
            except (ValueError, TypeError):
                vigencia_dt = None
        
        # Si no tenemos frecuencia, mantener el registro
        if not frecuencia:
            frecuencia = f"sin_frecuencia_{id(registro)}"
        
        # Si es la primera vez que vemos esta frecuencia, o si encontramos una vigencia más alta
        if frecuencia not in registros_por_frecuencia:
            registros_por_frecuencia[frecuencia] = registro
            registro['_VIGENCIA_DT'] = vigencia_dt
        else:
            registro_existente = registros_por_frecuencia[frecuencia]
            vigencia_existente = registro_existente.get('_VIGENCIA_DT')
            
            # Comparar vigencias
            if vigencia_dt and vigencia_existente:
                if vigencia_dt > vigencia_existente:
                    registros_por_frecuencia[frecuencia] = registro
                    registro['_VIGENCIA_DT'] = vigencia_dt
            elif vigencia_dt and not vigencia_existente:
                # Si el nuevo tiene vigencia y el existente no, usar el nuevo
                registros_por_frecuencia[frecuencia] = registro
                registro['_VIGENCIA_DT'] = vigencia_dt
            # Si ambos no tienen vigencia, mantener el primero
    
    # Crear lista final sin los campos internos
    resultados_finales = []
    for frecuencia, registro in registros_por_frecuencia.items():
        # Remover campos internos antes de guardar
        registro_final = {k: v for k, v in registro.items() if not k.startswith('_')}
        resultados_finales.append(registro_final)
    
    return resultados_finales

def mostrar_estadisticas(registros_iniciales, registros_finales):
    """
    Muestra estadísticas comparativas antes y después del filtro por vigencia
    """
    print(f"\n=== ESTADÍSTICAS ===")
    print(f"Registros iniciales después de filtros básicos: {len(registros_iniciales)}")
    print(f"Registros finales después de filtro por vigencia: {len(registros_finales)}")
    print(f"Registros eliminados por duplicados de frecuencia: {len(registros_iniciales) - len(registros_finales)}")
    
    # Estadísticas por servicio
    servicios_iniciales = {}
    servicios_finales = {}
    
    for registro in registros_iniciales:
        servicio = registro.get('RED', '')
        servicios_iniciales[servicio] = servicios_iniciales.get(servicio, 0) + 1
    
    for registro in registros_finales:
        servicio = registro.get('RED', '')
        servicios_finales[servicio] = servicios_finales.get(servicio, 0) + 1
    
    print(f"\nDistribución por servicios (INICIAL):")
    for servicio, cantidad in servicios_iniciales.items():
        print(f"  - {servicio}: {cantidad} registros")
    
    print(f"\nDistribución por servicios (FINAL):")
    for servicio, cantidad in servicios_finales.items():
        print(f"  - {servicio}: {cantidad} registros")
    
    # Mostrar ejemplos de frecuencias con múltiples registros
    frecuencias_duplicadas = {}
    for registro in registros_iniciales:
        frecuencia = registro.get('FRECUENCIA', 'sin_frecuencia')
        frecuencias_duplicadas[frecuencia] = frecuencias_duplicadas.get(frecuencia, 0) + 1
    
    frecuencias_con_duplicados = {freq: count for freq, count in frecuencias_duplicadas.items() if count > 1}
    
    if frecuencias_con_duplicados:
        print(f"\nFrecuencias con múltiples registros (antes del filtro):")
        for frecuencia, count in list(frecuencias_con_duplicados.items())[:5]:  # Mostrar solo las primeras 5
            print(f"  - {frecuencia}: {count} registros")
        
        # Mostrar un ejemplo detallado
        print(f"\nEjemplo de filtrado para una frecuencia:")
        frecuencia_ejemplo = list(frecuencias_con_duplicados.keys())[0]
        registros_ejemplo = [r for r in registros_iniciales if r.get('FRECUENCIA') == frecuencia_ejemplo]
        registro_elegido = [r for r in registros_finales if r.get('FRECUENCIA') == frecuencia_ejemplo]
        
        if registro_elegido:
            print(f"Frecuencia: {frecuencia_ejemplo}")
            print(f"Registros encontrados: {len(registros_ejemplo)}")
            print(f"Registro elegido (mayor vigencia):")
            print(f"  VIGENCIA: {registro_elegido[0].get('VIGENCIA', 'N/A')}")
            print(f"  NOMBRES: {registro_elegido[0].get('NOMBRES', 'N/A')}")
    
    print("\nPreview de los primeros registros finales:")
    for i, registro in enumerate(registros_finales[:3], 1):
        print(f"Registro {i}:")
        for key, value in registro.items():
            print(f"  {key}: {value}")
        print()

# Función alternativa con filtros más avanzados
def filtrar_datos_avanzado(archivo_excel, archivo_json, filtros_personalizados=None):
    """
    Versión más avanzada con múltiples opciones de filtrado para AREAS_OP_MR
    """
    try:
        workbook = openpyxl.load_workbook(archivo_excel)
        sheet = workbook['descarga']
        
        # Filtros por defecto
        filtros_default = {
            'PROVINCIA_A': 'ZAMORA CHINCHIPE',
            'SERVICIO': 'FM - Frecuencia Modulada',
            'AREAS_OP_MR': {
                'tipo': 'contiene',  # 'contiene', 'exacto', 'comienza_con', 'termina_con'
                'valor': 'ZAMORA',
                'case_sensitive': False
            }
        }
        
        # Combinar con filtros personalizados si se proporcionan
        if filtros_personalizados is None:
            filtros = filtros_default
        else:
            filtros = {**filtros_default, **filtros_personalizados}
        
        # Mapeo de columnas
        columnas_extraer = {
            'NOMBRES': 'C',
            'RED': 'E', 
            'FRECUENCIA': 'F',
            'ESTADO_ESTACION': 'AI',
            'ESTACION': 'AR',
            'ESTADO_SOLICITUD': 'BC',
            'SUSCRIPCION': 'AJ',
            'VIGENCIA': 'AK'
        }
        
        resultados = []
        
        print("Aplicando filtros avanzados...")
        
        for row in range(2, sheet.max_row + 1):
            cumple_filtros = True
            
            # Filtro PROVINCIA_A
            if 'PROVINCIA_A' in filtros and cumple_filtros:
                provincia = sheet[f'A{row}'].value
                if provincia != filtros['PROVINCIA_A']:
                    cumple_filtros = False
            
            # Filtro SERVICIO
            if 'SERVICIO' in filtros and cumple_filtros:
                servicio = sheet[f'B{row}'].value
                if servicio != filtros['SERVICIO']:
                    cumple_filtros = False
            
            # FILTRO AVANZADO AREAS_OP_MR
            if 'AREAS_OP_MR' in filtros and cumple_filtros:
                areas_op = sheet[f'I{row}'].value
                config_filtro = filtros['AREAS_OP_MR']
                
                if areas_op is None:
                    cumple_filtros = False
                else:
                    texto_buscar = str(areas_op)
                    texto_filtro = config_filtro['valor']
                    
                    # Aplicar diferentes tipos de filtro
                    if not config_filtro.get('case_sensitive', False):
                        texto_buscar = texto_buscar.upper()
                        texto_filtro = texto_filtro.upper()
                    
                    tipo_filtro = config_filtro.get('tipo', 'contiene')
                    
                    if tipo_filtro == 'contiene':
                        if texto_filtro not in texto_buscar:
                            cumple_filtros = False
                    elif tipo_filtro == 'exacto':
                        if texto_buscar != texto_filtro:
                            cumple_filtros = False
                    elif tipo_filtro == 'comienza_con':
                        if not texto_buscar.startswith(texto_filtro):
                            cumple_filtros = False
                    elif tipo_filtro == 'termina_con':
                        if not texto_buscar.endswith(texto_filtro):
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
                
                # Agregar el valor original de AREAS_OP_MR
                registro['AREAS_OP_MR_ORIGINAL'] = sheet[f'I{row}'].value
                resultados.append(registro)
        
        # Guardar en JSON
        with open(archivo_json, 'w', encoding='utf-8') as json_file:
            json.dump(resultados, json_file, ensure_ascii=False, indent=2)
        
        print(f"Proceso completado. Se encontraron {len(resultados)} registros.")
        
        # Mostrar resumen de valores AREAS_OP_MR encontrados
        if resultados:
            areas_encontradas = set(registro['AREAS_OP_MR_ORIGINAL'] for registro in resultados)
            print(f"\nValores de AREAS_OP_MR que cumplieron el filtro:")
            for area in sorted(areas_encontradas):
                print(f"  ✓ {area}")
        
        return resultados
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return []

# Función específica para buscar múltiples términos en AREAS_OP_MR
def filtrar_por_multiples_areas(archivo_excel, archivo_json, terminos_busqueda):
    """
    Filtra por múltiples términos en AREAS_OP_MR (OR lógico)
    """
    try:
        workbook = openpyxl.load_workbook(archivo_excel)
        sheet = workbook['descarga']
        
        # Convertir términos a mayúsculas para búsqueda case-insensitive
        terminos = [term.upper() for term in terminos_busqueda]
        
        resultados = []
        
        print(f"Buscando términos: {', '.join(terminos_busqueda)}")
        
        for row in range(2, sheet.max_row + 1):
            # Filtros básicos
            provincia = sheet[f'A{row}'].value
            servicio = sheet[f'B{row}'].value
            areas_op = sheet[f'I{row}'].value
            
            if (provincia == 'ZAMORA CHINCHIPE' and 
                servicio == 'FM - Frecuencia Modulada' and 
                areas_op is not None):
                
                # FILTRO MÚLTIPLE OPTIMIZADO - busca cualquier término
                areas_op_upper = str(areas_op).upper()
                cumple_filtro_areas = any(term in areas_op_upper for term in terminos)
                
                if cumple_filtro_areas:
                    registro = {
                        'NOMBRES': sheet[f'C{row}'].value or "",
                        'RED': sheet[f'E{row}'].value or "",
                        'FRECUENCIA': sheet[f'F{row}'].value or "",
                        'ESTADO_ESTACION': sheet[f'AI{row}'].value or "",
                        'ESTACION': sheet[f'AR{row}'].value or "",
                        'ESTADO_SOLICITUD': sheet[f'BC{row}'].value or "",
                        'SUSCRIPCION': sheet[f'AJ{row}'].value or "",
                        'VIGENCIA': sheet[f'AK{row}'].value or "",
                        'AREAS_OP_MR_ORIGINAL': areas_op
                    }
                    resultados.append(registro)
        
        # Guardar resultados
        with open(archivo_json, 'w', encoding='utf-8') as json_file:
            json.dump(resultados, json_file, ensure_ascii=False, indent=2)
        
        print(f"Encontrados {len(resultados)} registros con términos: {', '.join(terminos_busqueda)}")
        return resultados
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return []

if __name__ == "__main__":
    # Configuración
    archivo_excel = "9. SPECTRA_10sep2025.xlsx"
    archivo_json = "resultados_filtrados.json"
    
    # Opción 1: Filtro básico optimizado (RECOMENDADO)
    print("=== FILTRO BÁSICO OPTIMIZADO ===")
    resultados = filtrar_datos_arcotel(archivo_excel, archivo_json)
    
    # Opción 2: Filtro avanzado con más opciones
    # print("\n=== FILTRO AVANZADO ===")
    # filtros_avanzados = {
    #     'AREAS_OP_MR': {
    #         'tipo': 'contiene',  # 'contiene', 'exacto', 'comienza_con', 'termina_con'
    #         'valor': 'ZAMORA',
    #         'case_sensitive': False
    #     }
    # }
    # resultados = filtrar_datos_avanzado(archivo_excel, "resultados_avanzados.json", filtros_avanzados)
    
    # Opción 3: Búsqueda por múltiples términos
    # print("\n=== BÚSQUEDA MÚLTIPLE ===")
    # terminos = ['ZAMORA', 'YANTZAZA', 'CHINCHIPE']
    # resultados = filtrar_por_multiples_areas(archivo_excel, "resultados_multiples.json", terminos)