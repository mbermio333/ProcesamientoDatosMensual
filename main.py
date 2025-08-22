import pandas as pd
import calendar
import os

# ---------- CONFIGURACIÓN GENERAL ----------
ruta_entrada = "MedicionesCSV"
ruta_salida = "Promedios_mensuales"
os.makedirs(ruta_salida, exist_ok=True)

# ---------- PROCESAMIENTO DE CADA ARCHIVO ----------
for archivo_entrada in os.listdir(ruta_entrada):
    if archivo_entrada.endswith(".csv"):
        try:
            print(f"📄 Procesando archivo: {archivo_entrada}")

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

            # Nombre del mes (opcional, si lo quieres imprimir)
            nombre_mes_es = {
                1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
                5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
                9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
            }[mes_objetivo]

            # Filtrar por mes detectado
            df["Frecuencia (MHz)"] = df["Frecuencia (Hz)"] / 1_000_000
            df["Level (dBµV/m)"] = df["Level (dBµV/m)"].astype(str).str.replace(",", ".", regex=False).astype(float)
            df = df[df["Tiempo"].apply(lambda x: x.month == mes_objetivo)]

            if df.empty:
                print(f"⚠ No hay datos válidos para el mes más frecuente ({nombre_mes_es}) en {archivo_entrada}")
                continue

            df["DIA"] = pd.to_datetime(df["Tiempo"]).dt.day

            # Promedio diario por estación, frecuencia, día
            agrupado = df.groupby(["ESTACION", "Frecuencia (MHz)", "DIA"])["Level (dBµV/m)"].mean().reset_index()
            pivot = agrupado.pivot(index=["ESTACION", "Frecuencia (MHz)"], columns="DIA", values="Level (dBµV/m)")

            # Asegurar días del 1 al 31
            todos_los_dias = list(range(1, 32))
            pivot = pivot.reindex(columns=todos_los_dias, fill_value=0)

            # Promedio mensual de Level (dBµV/m), ignorando ceros
            promedios_sin_ceros = pivot.replace(0, pd.NA).mean(axis=1, skipna=True)
            pivot["Promedio (dBuV/m)"] = promedios_sin_ceros.astype(float).round(3)

            # Promedio mensual Bandwidth (Hz) → kHz
            bandwidth_promedios = (
                df.groupby(["ESTACION", "Frecuencia (MHz)"])["Bandwidth (Hz)"]
                .mean()
                .reset_index()
            )
            bandwidth_promedios["Promedio Ancho de Banda (KHz)"] = (bandwidth_promedios["Bandwidth (Hz)"] / 1000).round(3)
            bandwidth_promedios = bandwidth_promedios.drop(columns=["Bandwidth (Hz)"])
            pivot = pivot.merge(bandwidth_promedios, on=["ESTACION", "Frecuencia (MHz)"], how="left")

            # Ordenar por frecuencia
            pivot = pivot.sort_values(by="Frecuencia (MHz)")

            # Reemplazar ceros por guiones
            pivot[todos_los_dias] = pivot[todos_los_dias].replace(0, "-")

            # Agregar columnas vacías
            pivot["Medición Manual AB(KHz) o NIVEL (dBuV/m)"] = ""
            pivot["Observaciones"] = ""

            # Eliminar columna 'index' si existe
            pivot = pivot.reset_index(drop=True)  # Elimina el índice directamente

            # Generar nombre del archivo de salida
            base_nombre = archivo_entrada.split("_")[0]  # antes del primer guion bajo
            archivo_salida = f"{base_nombre}_promediostotal.csv"
            ruta_csv_salida = os.path.join(ruta_salida, archivo_salida)

            # Guardar archivo
            pivot.to_csv(ruta_csv_salida, index=False, float_format="%.3f")
            print(f"✅ Archivo generado: {ruta_csv_salida}")

        except Exception as e:
            print(f"❌ Error procesando {archivo_entrada}: {e}")
