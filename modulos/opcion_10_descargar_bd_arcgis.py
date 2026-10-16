from datetime import datetime
import pandas as pd
from modulos.comun import cargar_bd_arcgis, ruta_en_reportes

def ejecutar():
    print("\n--- TIPO DE BASE A DESCARGAR ---")
    print("1. BD Completa\n2. BD Resumen")
    opcion_bd = input("Seleccione una opción (1 o 2): ").strip()
    if opcion_bd not in ["1", "2"]:
        opcion_bd = "1"

    fecha_hoy = datetime.now().strftime("%y%m%d")
    nombre_archivo = f"{fecha_hoy}_BD_Riesgos_Resumen.xlsx" if opcion_bd == "2" else f"{fecha_hoy}_BD_Riesgos.xlsx"
    ruta_final = ruta_en_reportes(nombre_archivo, subcarpeta="BD_ArcGIS")

    columnas_a_borrar = [
        'globalid', 'CreationDate', 'Creator', 'EditDate', 'Editor',
        'direcci_n_del_gadpa', 'direcci_n_del_gadpa_other',
        'tipo_de_evento_other', 'field_50_other', 'SHAPE'
    ]

    try:
        sedf_completo = cargar_bd_arcgis()
        if sedf_completo is None:
            print("❌ No se pudo cargar la BD de ArcGIS en memoria.")
            return

        if not sedf_completo.empty:
            columnas_presentes = [c for c in columnas_a_borrar if c in sedf_completo.columns]
            sedf_completo = sedf_completo.drop(columns=columnas_presentes)

            if 'fecha_de_visita' in sedf_completo.columns:
                if pd.api.types.is_numeric_dtype(sedf_completo['fecha_de_visita']):
                    sedf_completo['fecha_de_visita'] = pd.to_datetime(sedf_completo['fecha_de_visita'], unit='ms', errors='coerce')
                else:
                    sedf_completo['fecha_de_visita'] = pd.to_datetime(sedf_completo['fecha_de_visita'], errors='coerce')
                sedf_completo = sedf_completo.sort_values(by='fecha_de_visita', ascending=False)
                sedf_completo['fecha_de_visita'] = sedf_completo['fecha_de_visita'].dt.strftime('%d/%m/%Y')

            if opcion_bd == "2":
                columnas_resumen = [
                    "objectid", "nombre_del_punto", "fecha_de_visita", "sector", "parroquia",
                    "cant_n", "tipo_de_evento", "fecha_y_hora_de_evento", "tipo_de_atenci_n",
                    "trailers", "tractores", "excavadoras", "cargadoras", "retro_excavadoras",
                    "volquetes", "cisternas", "drones", "cuadrilla", "operadores_y_conductores",
                    "personal_t_cnico", "en_intervenci_n"
                ]
                columnas_presentes_resumen = [c for c in columnas_resumen if c in sedf_completo.columns]
                sedf_salida = sedf_completo[columnas_presentes_resumen].copy()
            else:
                sedf_salida = sedf_completo

            sedf_salida.to_excel(ruta_final, index=False)
            print("-" * 45)
            print(f"✅ ¡Archivo guardado con éxito!\n📍 Ubicación: {ruta_final}")
            print("-" * 45)
        else:
            print("⚠️ No hay datos disponibles para descargar.")
    except Exception as e:
        print(f"❌ Ocurrió un error: {e}")