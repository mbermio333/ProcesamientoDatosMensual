import pandas as pd
import calendar
import os

# -------- CONFIGURACIÓN --------
# Ruta a la carpeta de entrada
ruta_entrada = "MedicionesCSV"
archivo_entrada = "canar_juliototal.csv"

# Ruta a la carpeta de salida
ruta_salida = "Promedios_estaciones"

# Asegurar que carpeta de salida exista
os.makedirs(ruta_salida, exist_ok=True)

# Construir rutas completas
ruta_csv_entrada = os.path.join(ruta_entrada, archivo_entrada)

# --------------------------------

# Cargar datos
df = pd.read_csv(ruta_csv_entrada, encoding='unicode_escape')
df = df.iloc[:, :9]
df["ESTACION"] = df["ESTACION"].astype(str).str.strip()
df = df[df["ESTACION"].notna() & (df["ESTACION"] != "")]

# Formatear columnas
df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date

# Detectar mes con más ocurrencias
mes_ocurrencias = df["Tiempo"].apply(lambda x: x.month)
mes_objetivo = mes_ocurrencias.value_counts().idxmax()

# Nombre del mes en español
nombre_mes_es = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}[mes_objetivo]

#ruta salida 
ruta_csv_salida = os.path.join(ruta_salida, f"promedios_{nombre_mes_es}_todas_estaciones.csv")

df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"] / 1_000_000
df["Level (dBµV/m)"] = df["Level (dBµV/m)"].astype(str).str.replace(",", ".", regex=False).astype(float)
df = df[df["Tiempo"].apply(lambda x: x.month == mes_objetivo)]

if df.empty:
    print(f"⚠ No hay datos para el mes {nombre_mes_es}")
else:
    df["DIA"] = pd.to_datetime(df["Tiempo"]).dt.day

    agrupado = df.groupby(["ESTACION", "Frecuencia (MHz)", "DIA"])["Level (dBµV/m)"].mean().reset_index()

    pivot = agrupado.pivot(index=["ESTACION", "Frecuencia (MHz)"], columns="DIA", values="Level (dBµV/m)")

    # Rellenar todos los días del mes con 0
    todos_los_dias = list(range(1, 32))
    pivot = pivot.reindex(columns=todos_los_dias, fill_value=0)

    # Promedio mensual (ignorando ceros)
    promedios_sin_ceros = pivot.replace(0, pd.NA).mean(axis=1, skipna=True)
    pivot["Promedio Mensual"] = promedios_sin_ceros.astype(float).round(3)

    # Ordenar por frecuencia
    pivot = pivot.sort_values(by="Frecuencia (MHz)")

    # Reemplazar ceros con '-'
    pivot[todos_los_dias] = pivot[todos_los_dias].replace(0, "-")

    pivot = pivot.reset_index()

    # Guardar resultado
    pivot.to_csv(ruta_csv_salida, index=False, float_format="%.3f")

    print(f"✅ Archivo generado en: {ruta_csv_salida}")