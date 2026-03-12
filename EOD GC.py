import os
import glob
import pandas as pd

# Ruta de Descargas
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
csv_files = glob.glob(os.path.join(downloads_path, "*.csv"))

if not csv_files:
    print("No se encontraron archivos CSV en Descargas.")
else:
    latest_csv = max(csv_files, key=os.path.getctime)
    print("Archivo seleccionado:", os.path.basename(latest_csv))

    # Leer el archivo con separador de punto y coma
    df = pd.read_csv(latest_csv, encoding="utf-8", sep=";")

    # Definir nombres de columnas
    columnas = df.columns.tolist()

    # ---- Filtros ----
    filtro1 = (df.iloc[:, 1] == 0) & (df.iloc[:, 42] != "ABANDON")   # Columna B y AQ
    filtro2 = (df.iloc[:, 23] == "7867362870")                       # Columna X
    filtro3 = df.iloc[:, 35].astype(str).str.lower().str.contains("coll", na=False)  # Columna 36

    # Contar eliminaciones específicas
    eliminadas_numero = df[filtro2].shape[0]
    eliminadas_coll = df[filtro3].shape[0]

    # Combinar filtros
    filtro_total = filtro1 | filtro2 | filtro3

    # Aplicar filtros
    df_filtrado = df.drop(df[filtro_total].index)

    # ---- Reporte ----
    print(f"Filas iniciales: {len(df)}")
    print(f"Filas eliminadas: {len(df) - len(df_filtrado)}")
    print(f"Llamadas finales: {len(df_filtrado)}")

    if eliminadas_numero > 0:
        print(f"Se eliminaron {eliminadas_numero} filas por contener el número 7867362870.")
    else:
        print("No se eliminaron filas por el número 7867362870.")

    if eliminadas_coll > 0:
        print(f"Se eliminaron {eliminadas_coll} filas por contener 'coll'.")
    else:
        print("No se eliminaron filas por 'coll'.")

    # ---- UTM ----
    utm_valores = ["Digital Discount", "PTC Standard", "Digital Standard", "PTC Discount"]
    utm_rows = df_filtrado[df_filtrado.iloc[:, 35].isin(utm_valores)]
    utm_volume = utm_rows.shape[0]
    print(f"UTM volume: {utm_volume}")

    # ---- Función para aplicar formatos ----
    def aplicar_formatos(df):
        df = df.copy()
        # Columna N (14 → índice 13)
        df[columnas[13]] = pd.to_datetime(df[columnas[13]], errors="coerce")
        # Columna S (19 → índice 18)
        df[columnas[18]] = pd.to_datetime(df[columnas[18]], errors="coerce")
        # Columna X (24 → índice 23)
        df[columnas[23]] = pd.to_numeric(df[columnas[23]], errors="coerce")
        return df

    # ---- Guardar archivo UTM ----
    if utm_volume > 0:
        utm_rows_fmt = aplicar_formatos(utm_rows)
        fecha = pd.to_datetime(utm_rows_fmt.iloc[0, 18], errors="coerce")
        if pd.notnull(fecha):
            fecha_str = fecha.strftime("%m-%d")
        else:
            fecha_str = "unknown"

        filename = f"UTM EOD {fecha_str}.xlsx"
        filepath = os.path.join(downloads_path, filename)

        utm_rows_fmt.to_excel(filepath, index=False, engine="openpyxl")
        print(f"Archivo UTM guardado en: {filepath}")
    else:
        print("No se encontraron filas UTM para guardar.")

    # ---- Guardar llamadas finales ----
    if len(df_filtrado) > 0:
        df_llamadas_fmt = aplicar_formatos(df_filtrado)
        fecha_llamadas = pd.to_datetime(df_llamadas_fmt.iloc[0, 18], errors="coerce")
        if pd.notnull(fecha_llamadas):
            fecha_llamadas_str = fecha_llamadas.strftime("%m%d")
        else:
            fecha_llamadas_str = "unknown"

        filename_llamadas = f"llamadas reales {fecha_llamadas_str}.xlsx"
        filepath_llamadas = os.path.join(downloads_path, filename_llamadas)

        df_llamadas_fmt.to_excel(filepath_llamadas, index=False, engine="openpyxl")
        print(f"Archivo de llamadas reales guardado en: {filepath_llamadas}")
    else:
        print("No hay llamadas finales para guardar.")
