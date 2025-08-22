import pandas as pd
import calendar
import os

# -------- CONFIGURACIÓN --------
ruta_entrada = "MedicionesCSV"
archivo_entrada = "canar_juliototal.csv"  # ⚠️ Cambia el nombre si usas otro archivo
ruta_salida = "pruebas"
os.makedirs(ruta_salida, exist_ok=True)

# Construir rutas
ruta_csv_entrada = os.path.join(ruta_entrada, archivo_entrada)

# Cargar datos
df = pd.read_csv(ruta_csv_entrada, encoding='unicode_escape')
df = df.iloc[:, :9]
df["ESTACION"] = df["ESTACION"].astype(str).str.strip()
df = df[df["ESTACION"].notna() & (df["ESTACION"] != "")]

# Formatear columnas
df["Tiempo"] = pd.to_datetime(df["Tiempo"], format="%d/%m/%Y  %H:%M:%S,%f", errors='coerce').dt.date

# Detectar mes con más datos
mes_ocurrencias = df["Tiempo"].apply(lambda x: x.month)
mes_objetivo = mes_ocurrencias.value_counts().idxmax()

nombre_mes_es = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}[mes_objetivo]

# Filtrar por mes
df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"] / 1_000_000
df["Level (dBµV/m)"] = df["Level (dBµV/m)"].astype(str).str.replace(",", ".", regex=False).astype(float)
df = df[df["Tiempo"].apply(lambda x: x.month == mes_objetivo)]

if df.empty:
    print(f"⚠ No hay datos válidos para el mes más frecuente ({nombre_mes_es}) en {archivo_entrada}")
else:
    df["DIA"] = pd.to_datetime(df["Tiempo"]).dt.day

    agrupado = df.groupby(["ESTACION", "Frecuencia (MHz)", "DIA"])["Level (dBµV/m)"].mean().reset_index()
    pivot = agrupado.pivot(index=["ESTACION", "Frecuencia (MHz)"], columns="DIA", values="Level (dBµV/m)")

    # Rellenar días del 1 al 31
    todos_los_dias = list(range(1, 32))
    pivot = pivot.reindex(columns=todos_los_dias, fill_value=0)

    # Promedio mensual Level
    promedios_sin_ceros = pivot.replace(0, pd.NA).mean(axis=1, skipna=True)
    pivot["Promedio Mensual"] = promedios_sin_ceros.astype(float).round(3)

    # Promedio mensual Bandwidth en kHz
    bandwidth_promedios = (
        df.groupby(["ESTACION", "Frecuencia (MHz)"])["Bandwidth (Hz)"]
        .mean()
        .reset_index()
    )
    bandwidth_promedios["Promedio Bandwidth (kHz)"] = (bandwidth_promedios["Bandwidth (Hz)"] / 1000).round(3)
    bandwidth_promedios = bandwidth_promedios.drop(columns=["Bandwidth (Hz)"])
    pivot = pivot.merge(bandwidth_promedios, on=["ESTACION", "Frecuencia (MHz)"], how="left")

    # Ordenar
    pivot = pivot.sort_values(by="Frecuencia (MHz)")

    # Reemplazar ceros con guiones
    pivot[todos_los_dias] = pivot[todos_los_dias].replace(0, "-")

    # Columnas manuales vacías
    pivot["Medicion Manual"] = ""
    pivot["Observaciones"] = ""

    # Resetear índice y eliminar columna 'index' si aparece
    pivot = pivot.reset_index(drop=True)

    # Generar nombre del archivo de salida
    base_nombre = archivo_entrada.split("_")[0]
    nombre_excel = f"{base_nombre}_promediostotal.xlsx"
    ruta_excel_salida = os.path.join(ruta_salida, nombre_excel)

    # Guardar en formato Excel
    with pd.ExcelWriter(ruta_excel_salida, engine="openpyxl") as writer:
        pivot.to_excel(writer, index=False, sheet_name="Resumen")

    print(f"✅ Archivo generado en formato Excel: {ruta_excel_salida}")
