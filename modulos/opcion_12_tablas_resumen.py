import warnings
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from modulos.comun import cargar_bd_arcgis, ruta_en_reportes

def ejecutar():
    warnings.filterwarnings("ignore")
    print("\n--- OPCIONES DE FILTRADO ---")
    print("1. Filtrar solo el año 2025\n2. Filtrar solo el año 2026\n3. Rango de fechas personalizado")
    print("4. Frentes Activos (Solo en intervención: 'Sí')\n5. Solo emergencias (Filtro por columna tipo_de_atenci_n)")
    opcion_filtro = input("Elige una opción (1, 2, 3, 4 o 5): ").strip()

    fecha_inicio_dt = None
    fecha_fin_dt = None

    if opcion_filtro == "3":
        fecha_inicio_str = input("Ingresa la fecha de INICIO (formato DD/MM/AAAA): ").strip()
        fecha_fin_str = input("Ingresa la fecha de FIN (formato DD/MM/AAAA): ").strip()
        fecha_inicio_dt = pd.to_datetime(fecha_inicio_str, format="%d/%m/%Y", errors="coerce")
        fecha_fin_dt = pd.to_datetime(fecha_fin_str, format="%d/%m/%Y", errors="coerce")
        if pd.isna(fecha_inicio_dt) or pd.isna(fecha_fin_dt):
            opcion_filtro = "2"
    elif opcion_filtro not in ["1", "2", "3", "4", "5"]:
        opcion_filtro = "2"

    try:
        df = cargar_bd_arcgis()
        if df is None or df.empty:
            print("❌ No se pudo cargar la BD de ArcGIS (o está vacía).")
            return
        df = df.copy()

        if "fecha_de_visita" in df.columns:
            if pd.api.types.is_numeric_dtype(df["fecha_de_visita"]):
                df["fecha_de_visita"] = pd.to_datetime(df["fecha_de_visita"], unit="ms", errors="coerce")
            else:
                df["fecha_de_visita"] = pd.to_datetime(df["fecha_de_visita"], errors="coerce")

        df_filtrado = df.copy()
        mascara = None
        if opcion_filtro == "1":
            mascara = df_filtrado["fecha_de_visita"].dt.year == 2025
        elif opcion_filtro == "2":
            mascara = df_filtrado["fecha_de_visita"].dt.year == 2026
        elif opcion_filtro == "3":
            mascara = (df_filtrado["fecha_de_visita"] >= fecha_inicio_dt) & (
                df_filtrado["fecha_de_visita"] <= fecha_fin_dt + pd.Timedelta(days=1, seconds=-1)
            )
        elif opcion_filtro == "4":
            if "en_intervenci_n" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["en_intervenci_n"].astype(str).str.strip().str.lower().isin(["sí", "si"])]
            mascara = None
        elif opcion_filtro == "5":
            if "tipo_de_atenci_n" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["tipo_de_atenci_n"].astype(str).str.strip().str.lower() == "emergencia"]
            mascara = None

        if opcion_filtro in ["1", "2", "3"] and mascara is not None:
            if "en_intervenci_n" in df_filtrado.columns:
                df_filtrado = df_filtrado[mascara | df_filtrado["en_intervenci_n"].astype(str).str.strip().str.lower().isin(["sí", "si"])]
            else:
                df_filtrado = df_filtrado[mascara]

        columnas_equipos = {
            "trailers": ("Trailer", "Trailers"), "tractores": ("Tractor", "Tractores"),
            "excavadoras": ("Excavadora", "Excavadoras"), "cargadoras": ("Cargadora", "Cargadoras"),
            "retro_excavadoras": ("Retroexcavadora", "Retroexcavadoras"), "mini_cargadoras": ("Minicargadora", "Minicargadoras"),
            "volquetes": ("Volquete", "Volquetes"), "cisternas": ("Cisterna", "Cisternas"),
            "camionetas": ("Camioneta", "Camionetas"), "drones": ("Dron", "Drones")
        }

        def juntar_equipos(row):
            equipos = []
            for col, (singular, plural) in columnas_equipos.items():
                if col in row and pd.notna(row[col]):
                    try:
                        cant = int(float(row[col]))
                        if cant == 1: equipos.append(f"1 {singular}")
                        elif cant > 1: equipos.append(f"{cant} {plural}")
                    except Exception: pass
            return ", ".join(equipos) if equipos else "Sin equipos"

        df_filtrado["equipos"] = df_filtrado.apply(juntar_equipos, axis=1)

        columnas_personal = {
            "cuadrilla": ("Cuadrilla", "Cuadrillas"),
            "operadores_y_conductores": ("Operador", "Operadores"),
            "personal_t_cnico": ("Técnico", "Técnicos")
        }

        def juntar_personal(row):
            personal = []
            for col, (singular, plural) in columnas_personal.items():
                if col in row and pd.notna(row[col]):
                    try:
                        cant = int(float(row[col]))
                        if cant == 1: personal.append(f"1 {singular}")
                        elif cant > 1: personal.append(f"{cant} {plural}")
                    except Exception: pass
            return ", ".join(personal) if personal else "Sin personal"

        df_filtrado["personal"] = df_filtrado.apply(juntar_personal, axis=1)

        columnas_finales = [
            col for col in [
                "nombre_del_punto", "parroquia", "cant_n", "fecha_de_visita",
                "subtipo_de_evento", "equipos", "personal"
            ] if col in df_filtrado.columns
        ]
        if "en_intervenci_n" in df_filtrado.columns: columnas_finales.append("en_intervenci_n")
        if "descripci_n_intervenci_n" in df_filtrado.columns: columnas_finales.append("descripci_n_intervenci_n")

        df_resultado = df_filtrado[columnas_finales].copy()
        if "fecha_de_visita" in df_resultado.columns:
            df_resultado = df_resultado.sort_values(by="fecha_de_visita", ascending=True)
            df_resultado["fecha_de_visita"] = df_resultado["fecha_de_visita"].dt.strftime("%d/%m/%Y")

        df_resultado.rename(columns={
            "nombre_del_punto": "Sector", "parroquia": "Parroquia", "cant_n": "Cantón",
            "fecha_de_visita": "Fecha de visita", "subtipo_de_evento": "Subtipo de evento",
            "en_intervenci_n": "En intervención", "descripci_n_intervenci_n": "Actividades realizadas"
        }, inplace=True)
        df_resultado.insert(0, "No.", range(1, len(df_resultado) + 1))

        if "cant_n" in df_filtrado.columns and "parroquia" in df_filtrado.columns:
            df_atenciones = df_filtrado.groupby(["cant_n", "parroquia"]).size().reset_index(name="Total Atenciones")
        else:
            df_atenciones = pd.DataFrame()

        cols_a_sumar = list(columnas_equipos.keys()) + list(columnas_personal.keys())
        for col in cols_a_sumar:
            if col in df_filtrado.columns:
                df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce").fillna(0)
            else:
                df_filtrado[col] = 0

        totales = df_filtrado[cols_a_sumar].sum()
        df_totales = pd.DataFrame([
            {"Recursos": "Trailers", "Cantidad": totales.get("trailers", 0)},
            {"Recursos": "Tractores", "Cantidad": totales.get("tractores", 0)},
            {"Recursos": "Excavadoras", "Cantidad": totales.get("excavadoras", 0)},
            {"Recursos": "Cargadoras", "Cantidad": totales.get("cargadoras", 0)},
            {"Recursos": "Retroexcavadoras", "Cantidad": totales.get("retro_excavadoras", 0)},
            {"Recursos": "Volquetes", "Cantidad": totales.get("volquetes", 0)},
            {"Recursos": "Cisternas", "Cantidad": totales.get("cisternas", 0)},
            {"Recursos": "Camionetas", "Cantidad": totales.get("camionetas", 0)},
            {"Recursos": "No. vuelo de drones", "Cantidad": totales.get("drones", 0)},
            {"Recursos": "Personal Operativo", "Cantidad": totales.get("operadores_y_conductores", 0) + totales.get("cuadrilla", 0)},
            {"Recursos": "Personal Técnico", "Cantidad": totales.get("personal_t_cnico", 0)},
        ])

        fecha_ejecucion_str = datetime.now().strftime("%y%m%d")
        if opcion_filtro == "1": nombre_filtro = "Año2025"
        elif opcion_filtro == "2": nombre_filtro = "Año2026"
        elif opcion_filtro == "3":
            fi_str = fecha_inicio_dt.strftime("%y%m%d") if fecha_inicio_dt is not None else "000000"
            ff_str = fecha_fin_dt.strftime("%y%m%d") if fecha_fin_dt is not None else "000000"
            nombre_filtro = f"Rango{fi_str}-{ff_str}"
        elif opcion_filtro == "4": nombre_filtro = "FrentesActivos"
        elif opcion_filtro == "5": nombre_filtro = "SoloEmergencias"
        else: nombre_filtro = "SinFiltro"

        ruta_salida = str(ruta_en_reportes(f"{fecha_ejecucion_str}-{nombre_filtro}.xlsx", subcarpeta="TablasResumen"))
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, sheet_name="Resumen General", index=False)
            df_atenciones.to_excel(writer, sheet_name="Atenciones por Sector", index=False)
            df_totales.to_excel(writer, sheet_name="Total Equipos", index=False)

        wb = load_workbook(ruta_salida)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 25

        wb.save(ruta_salida)
        print(f"\n¡PROCESO COMPLETADO! Archivo generado en: {ruta_salida}")
    except Exception as e:
        print(f"\nError al generar las tablas de resumen: {e}")