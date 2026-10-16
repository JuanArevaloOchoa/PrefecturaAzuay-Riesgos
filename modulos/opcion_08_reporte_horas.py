import os
from pathlib import Path
import pandas as pd
from modulos.comun import ruta_en_reportes

def ejecutar():
    print("--- PROCESADOR DE ASISTENCIA ---")
    ruta_archivo = input("Introduce la ruta completa del archivo (ej: C:/Carpeta/asistencia.xlsx): ").strip()
    ruta_archivo = ruta_archivo.replace('"', '')

    if not os.path.exists(ruta_archivo):
        print(f"Error: No se encontró el archivo en {ruta_archivo}")
        return

    try:
        if ruta_archivo.lower().endswith('.csv'):
            try:
                df_raw = pd.read_csv(ruta_archivo, header=None, encoding='utf-8')
            except Exception:
                df_raw = pd.read_csv(ruta_archivo, header=None, encoding='latin1')
        else:
            df_raw = pd.read_excel(ruta_archivo, header=None)

        idx_header = df_raw[df_raw.astype(str).apply(lambda x: x.str.contains('Fecha', case=False)).any(axis=1)].index[0]
        df = df_raw.iloc[idx_header:].copy()
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

        df = df.dropna(subset=['Fecha', 'Hora'])
        df['Fecha_dt'] = pd.to_datetime(df['Fecha'], dayfirst=True)
        df = df.drop_duplicates(subset=['Fecha_dt', 'Hora'])
        df = df.sort_values(by=['Fecha_dt', 'Hora'], ascending=[True, True])
        df['Num'] = df.groupby('Fecha_dt').cumcount() + 1
        df = df[df['Num'] <= 4]

        reporte = df.pivot(index='Fecha_dt', columns='Num', values='Hora')
        columnas_finales = {1: 'ENTRADA', 2: 'SALIDA ALMUERZO', 3: 'ENTRADA ALMUERZO', 4: 'SALIDA'}
        reporte = reporte.rename(columns=columnas_finales)
        reporte = reporte.sort_index(ascending=False)

        nombre_salida = f"{Path(ruta_archivo).stem}_REPORTE_FORMATO_FINAL.xlsx"
        ruta_salida = str(ruta_en_reportes(nombre_salida, subcarpeta="Asistencia"))

        with pd.ExcelWriter(ruta_salida, engine='openpyxl', date_format='DD/MM/YYYY') as writer:
            reporte.to_excel(writer, sheet_name='Asistencia')

        print(f"\n✅ Proceso completado con éxito.\n📂 Archivo guardado en: {ruta_salida}")
    except Exception as e:
        print(f"\n❌ Ocurrió un error: {e}\nRevisa que el archivo no esté abierto en Excel.")