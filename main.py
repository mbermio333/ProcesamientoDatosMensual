import pandas as pd
import calendar

# Leer el archivo original
df = pd.read_csv("/Users/mateobermeo/Desktop/ProcesamientoDatosMensual/Mediciones_totalesCSV/FmL030_juliototal.csv", encoding='unicode_escape')

# Conservar solo las primeras 9 columnas
df = df.iloc[:, :9]

# Limpiar columna ESTACION
df["ESTACION"] = df["ESTACION"].astype(str).str.strip()
df = df[df["ESTACION"].notna() & (df["ESTACION"] != "")]

# Convertir Tiempo a fecha
df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date

# Convertir frecuencia a MHz
df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"] / 1_000_000

# Convertir "Level (dBµV/m)" a float
df["Level (dBµV/m)"] = (
    df["Level (dBµV/m)"]
    .astype(str)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

# Establecer el mes a filtrar (ej: julio = 7)
mes_objetivo = 7

# Filtrar solo ese mes
df = df[df["Tiempo"].apply(lambda x: x.month == mes_objetivo)]

# Si el dataframe queda vacío, abortamos
if df.empty:
    print(f"⚠ No hay datos para el mes {calendar.month_name[mes_objetivo]}")
else:
    # Extraer día del mes
    df["DIA"] = pd.to_datetime(df["Tiempo"]).dt.day
    # Agrupar por estación, frecuencia y día
    agrupado = (
        df.groupby(["ESTACION", "Frecuencia (MHz)", "DIA"])["Level (dBµV/m)"]
        .mean()
        .reset_index()
    )

    # Pivot: fila por estación/frecuencia, columnas por día
    pivot = agrupado.pivot(index=["ESTACION", "Frecuencia (MHz)"], columns="DIA", values="Level (dBµV/m)")

    # Asegurar todos los días del mes (1 al 31), rellenar con 0
    todos_los_dias = list(range(1, 32))
    pivot = pivot.reindex(columns=todos_los_dias, fill_value=0)

   # Calcular promedio mensual ignorando ceros
    promedios_sin_ceros = pivot.replace(0, pd.NA).mean(axis=1, skipna=True)

# Convertir a float y redondear a 3 decimales
    pivot["Promedio Mensual"] = promedios_sin_ceros.astype(float).round(3)

    # 1. Ordenar por frecuencia
    pivot = pivot.sort_values(by="Frecuencia (MHz)")

    # 2. Reemplazar ceros por '-' en las columnas de días
    pivot[todos_los_dias] = pivot[todos_los_dias].replace(0, "-")


    # Resetear índice
    pivot = pivot.reset_index()

    # Generar nombre del archivo con el nombre del mes (en minúsculas)
    nombre_mes_es = {
         1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
         5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
         9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }[mes_objetivo]
    
    ruta_salida = "/Users/mateobermeo/Desktop/ProcesamientoDatosMensual/Promedios_estaciones/"
    nombre_archivo = f"{ruta_salida}promedios_{nombre_mes_es}_todas_estaciones.csv"


    # Guardar archivo con 3 decimales
    pivot.to_csv(nombre_archivo, index=False, float_format="%.3f")

   
    print(f"✅ Archivo generado: {nombre_archivo}")
