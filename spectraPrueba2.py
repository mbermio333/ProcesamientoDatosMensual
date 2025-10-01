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
        
        # Definir los filtros (AHORA MÁS FLEXIBLES)
        filtros = {
            'PROVINCIA_A': 'CAÑAR',
            'SERVICIO': 'FM - Frecuencia Modulada', 
            'AREAS_OP_MR': 'TAMBO'  # Ahora busca cualquier campo que CONTENGA "ZAMORA"
        }
        
        # Mapeo de columnas a extraer
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
        
        # Obtener índices de las columnas de filtro
        col_provincia = 'A'  # PROVINCIA_A
        col_servicio = 'B'   # SERVICIO
        col_areas_op = 'I'   # AREAS_OP_MR
        
        # Lista para almacenar los resultados
        resultados = []
        
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
            
            # Aplicar filtros (FILTRO AREAS_OP_MR ACTUALIZADO)
            cumple_filtros = True
            
            # Filtro PROVINCIA_A (exacto)
            if provincia != filtros['PROVINCIA_A']:
                cumple_filtros = False
            
            # Filtro SERVICIO (exacto)
            if servicio != filtros['SERVICIO']:
                cumple_filtros = False
            
            # FILTRO AREAS_OP_MR OPTIMIZADO - busca contenido en lugar de coincidencia exacta
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
                resultados.append(registro)
        
        # Guardar resultados en archivo JSON
        with open(archivo_json, 'w', encoding='utf-8') as json_file:
            json.dump(resultados, json_file, ensure_ascii=False, indent=2)
        
        print(f"\nProceso completado. Se encontraron {len(resultados)} registros.")
        print(f"Resultados guardados en: {archivo_json}")
        
        # Mostrar estadísticas de AREAS_OP_MR encontradas
        if resultados:
            areas_unicas = set(registro['_FILTRO_AREAS_OP_MR'] for registro in resultados)
            print(f"\nValores únicos de AREAS_OP_MR encontrados:")
            for area in sorted(areas_unicas):
                print(f"  - {area}")
            
            print("\nPreview de los primeros registros:")
            for i, registro in enumerate(resultados[:3], 1):
                print(f"Registro {i}:")
                for key, value in registro.items():
                    if not key.startswith('_'):  # No mostrar campos internos
                        print(f"  {key}: {value}")
                print()
        
        return resultados
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {archivo_excel}")
        return []
    except Exception as e:
        print(f"Error durante el procesamiento: {str(e)}")
        return []

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